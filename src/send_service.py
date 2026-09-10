"""步骤四：校验对账报告 + 组装邮件 + SMTP 发送。"""

from __future__ import annotations

import logging
import re
from dataclasses import dataclass
from pathlib import Path

from openpyxl import load_workbook

from .config_loader import load_config
from .local_settings import load_local_email
from .mailer import MailPayload, resolve_smtp, send_smtp, total_attachment_mb
from .paths import app_root

logger = logging.getLogger(__name__)


@dataclass
class SendPlan:
    bill_name: str
    pdf_dir: Path
    report_path: Path | None
    pdfs: list[Path]
    subject: str
    body: str
    to: list[str]
    cc: list[str]
    batch_ok: bool
    message: str
    attachment_mb: float


def extract_bill_name(text: str) -> str:
    m = re.search(r"(#DS-\d+-SMART)", text, re.I)
    if m:
        return m.group(1)
    stem = Path(text).name
    m2 = re.search(r"(#DS-\d+-SMART)", stem, re.I)
    return m2.group(1) if m2 else Path(stem).stem


def _read_report_status(report_path: Path) -> tuple[bool, str, str | None, int]:
    if not report_path.exists():
        return False, f"缺少对账报告: {report_path}", None, 0
    wb = load_workbook(report_path, data_only=True)
    ws = wb.active
    bill_name = None
    batch_ok = False
    message = ""
    ok_count = 0
    for r in range(2, (ws.max_row or 1) + 1):
        status = str(ws.cell(r, 6).value or "").strip().upper()
        if status in {"OK", "WARN"}:
            ok_count += 1
        a = ws.cell(r, 1).value
        b = ws.cell(r, 2).value
        if a == "账单名称":
            bill_name = str(b) if b else None
        if a == "批次是否放行":
            batch_ok = str(b).strip().upper() in {"YES", "TRUE", "1", "是"}
        if a == "批次说明":
            message = str(b or "")
    return batch_ok, message, bill_name, ok_count


def collect_pdfs(pdf_dir: Path) -> list[Path]:
    return sorted(
        p for p in Path(pdf_dir).iterdir() if p.is_file() and p.suffix.lower() == ".pdf"
    )


def find_report_for_pdf_dir(pdf_dir: Path) -> Path:
    pdf_dir = Path(pdf_dir)
    root = app_root()
    candidates = [
        root / "logs" / pdf_dir.name / "report.xlsx",
        pdf_dir.parent.parent / "logs" / pdf_dir.name / "report.xlsx",
        root / "logs" / "report.xlsx",
    ]
    for c in candidates:
        if c.exists():
            return c
    return candidates[0]


def build_send_plan(pdf_dir: Path, cfg: dict | None = None) -> SendPlan:
    cfg = cfg or load_config()
    email_cfg = cfg.get("email") or {}
    pdf_dir = Path(pdf_dir)
    report_path = find_report_for_pdf_dir(pdf_dir)
    pdfs = collect_pdfs(pdf_dir)
    batch_ok, msg, bill_from_report, ok_count = _read_report_status(report_path)
    bill_name = bill_from_report or extract_bill_name(pdf_dir.name)

    # 注意：cc=[] 表示明确不抄送，不能用 `or`（空列表在 Python 为假）
    to = list(email_cfg["to"]) if "to" in email_cfg else ["ACCT19.BD@smartsupplychaininc.com"]
    cc = list(email_cfg["cc"]) if "cc" in email_cfg else ["tracy@tofba.com", "eric.fan@tofba.com"]
    sender_name = email_cfg.get("sender_name") or "Raelyn.Li"
    subject_tpl = email_cfg.get("subject_template") or "{bill_name}账单invoice"
    body_tpl = email_cfg.get("body_template") or (
        "Hello, SMART team.\n\n"
        "{bill_name}账单invoice已附上，请查收，谢谢！\n\n"
        "Best Regards,\n"
        "{sender_name}"
    )
    subject = subject_tpl.format(bill_name=bill_name, sender_name=sender_name)
    body = body_tpl.format(bill_name=bill_name, sender_name=sender_name)

    issues: list[str] = []
    if not report_path.exists():
        if email_cfg.get("require_batch_ok", True):
            issues.append(f"缺少对账报告: {report_path}")
    elif email_cfg.get("require_batch_ok", True) and not batch_ok:
        issues.append(f"对账未放行: {msg}")
    if not pdfs:
        issues.append("没有可发送的 PDF")
    if (
        email_cfg.get("require_pdf_count_match", True)
        and report_path.exists()
        and ok_count
        and len(pdfs) != ok_count
    ):
        issues.append(f"PDF 数量({len(pdfs)})与报告成功柜数({ok_count})不一致")

    ok = len(issues) == 0
    return SendPlan(
        bill_name=bill_name,
        pdf_dir=pdf_dir,
        report_path=report_path if report_path.exists() else None,
        pdfs=pdfs,
        subject=subject,
        body=body,
        to=to,
        cc=cc,
        batch_ok=ok,
        message="; ".join(issues) if issues else (msg or "OK"),
        attachment_mb=total_attachment_mb(pdfs) if pdfs else 0.0,
    )


def send_plan(plan: SendPlan, cfg: dict | None = None) -> None:
    if not plan.batch_ok:
        raise RuntimeError(f"发送被阻断: {plan.message}")
    cfg = cfg or load_config()
    email_cfg = cfg.get("email") or {}
    local = load_local_email()
    smtp = resolve_smtp(email_cfg, local)
    max_mb = float(email_cfg.get("max_attachment_mb") or 20)

    pdfs = plan.pdfs
    if plan.attachment_mb > max_mb and email_cfg.get("split_when_too_large", True):
        chunks: list[list[Path]] = []
        current: list[Path] = []
        size = 0.0
        for p in pdfs:
            mb = p.stat().st_size / (1024 * 1024)
            if current and size + mb > max_mb:
                chunks.append(current)
                current, size = [], 0.0
            current.append(p)
            size += mb
        if current:
            chunks.append(current)
    elif plan.attachment_mb > max_mb:
        raise RuntimeError(f"附件约 {plan.attachment_mb:.1f}MB，超过限制 {max_mb}MB")
    else:
        chunks = [pdfs]

    from_addr = local.get("username") or email_cfg.get("username") or smtp.username
    for i, chunk in enumerate(chunks, 1):
        subject = plan.subject
        body = plan.body
        if len(chunks) > 1:
            subject = f"{plan.subject} (Part {i}/{len(chunks)})"
            body = f"{plan.body}\n\n（附件分多封发送：第 {i}/{len(chunks)} 封）"
        payload = MailPayload(
            subject=subject,
            body=body,
            to=plan.to,
            cc=plan.cc,
            attachments=chunk,
            from_addr=from_addr,
        )
        logger.info(
            "Sending mail part %s/%s attachments=%s to=%s",
            i,
            len(chunks),
            len(chunk),
            plan.to,
        )
        send_smtp(smtp, payload)
