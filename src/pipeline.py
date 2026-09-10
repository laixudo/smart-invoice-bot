from __future__ import annotations

import logging
from pathlib import Path

from .config_loader import load_config
from .excel_parser import parse_excel
from .excel_template_fill import fill_invoice_xlsx
from .excel_to_pdf import OfficeExporter, excel_export_pdf
from .pdf_renderer import render_invoice_pdf
from .report_writer import write_report
from .validator import validate_batch

logger = logging.getLogger(__name__)


def _render_one(
    invoice,
    out_pdf: Path,
    work_dir: Path,
    cfg: dict,
    mode: str,
    office_engine: str = "auto",
    exporter: OfficeExporter | None = None,
) -> Path:
    """
    mode:
      - template: 填 Excel 模板 → Excel/WPS 导出 PDF（推荐，对齐黄金样例）
      - reportlab: 纯代码绘制 PDF（无办公软件时的兜底）
    """
    if mode == "reportlab":
        return render_invoice_pdf(invoice, out_pdf, cfg)

    xlsx_path = work_dir / f"{invoice.container_no}.xlsx"
    tpl = cfg.get("invoice_template_xlsx")
    fill_invoice_xlsx(invoice, xlsx_path, Path(tpl) if tpl else None)
    try:
        return excel_export_pdf(
            xlsx_path, out_pdf, engine=office_engine, exporter=exporter
        )
    except Exception as e:
        logger.warning("办公软件导出 PDF 失败，回退 ReportLab: %s", e)
        return render_invoice_pdf(invoice, out_pdf, cfg)


def run_pipeline(
    input_xlsx: str | Path,
    output_dir: str | Path,
    config_path: str | Path | None = None,
    logs_dir: str | Path | None = None,
    render_mode: str | None = None,
    office_engine: str = "auto",
    cancel_check=None,
) -> dict:
    cfg = load_config(config_path)
    mode = render_mode or cfg.get("render_mode", "template")

    input_xlsx = Path(input_xlsx)
    output_dir = Path(output_dir)
    work_dir = Path(output_dir).resolve().parent / "_work" / "filled_xlsx" / Path(output_dir).name
    if logs_dir is None:
        logs_dir = Path(output_dir).resolve().parent / "_logs" / Path(output_dir).name
    logs_dir = Path(logs_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    logs_dir.mkdir(parents=True, exist_ok=True)
    work_dir.mkdir(parents=True, exist_ok=True)

    if cancel_check and cancel_check():
        raise InterruptedError("用户已中断")

    logger.info("Parsing %s", input_xlsx)
    parsed = parse_excel(input_xlsx, cfg)
    logger.info(
        "Parsed %d containers, %d quarantine rows, bill=%s date=%s",
        len(parsed.invoices),
        len(parsed.quarantine),
        parsed.bill_name,
        parsed.invoice_date,
    )

    if cancel_check and cancel_check():
        raise InterruptedError("用户已中断")

    batch = validate_batch(parsed, cfg)
    report_path = write_report(batch, logs_dir / "report.xlsx")
    logger.info("Validation batch_ok=%s msg=%s mode=%s", batch.batch_ok, batch.batch_message, mode)

    generated: list[str] = []
    if not batch.batch_ok:
        logger.error("Batch blocked. No PDFs written. See %s", report_path)
        return {
            "batch_ok": False,
            "message": batch.batch_message,
            "report": str(report_path),
            "generated": [],
            "container_count": len(batch.invoices),
            "bill_name": batch.bill_name,
            "system_total": str(batch.system_grand_total),
            "summary_total": str(batch.summary_grand_total),
            "render_mode": mode,
            "cancelled": False,
        }

    fail_set = {v.container_no for v in batch.validations if v.status == "FAIL"}
    to_make = [inv for inv in batch.invoices if inv.container_no not in fail_set]

    def _export_all(exporter: OfficeExporter | None) -> bool:
        """返回是否被用户中断。"""
        for inv in to_make:
            if cancel_check and cancel_check():
                return True
            pattern = cfg.get("output_name_pattern", "{container_no}.pdf")
            name = pattern.format(container_no=inv.container_no)
            out = output_dir / name
            _render_one(
                inv,
                out,
                work_dir,
                cfg,
                mode,
                office_engine=office_engine,
                exporter=exporter,
            )
            generated.append(str(out))
            logger.info("Wrote %s total=%s", out.name, inv.total_amount)
        return False

    cancelled = False
    if mode == "template" and to_make:
        try:
            with OfficeExporter(office_engine) as exporter:
                cancelled = _export_all(exporter)
        except Exception as e:
            logger.warning("批量办公导出会话失败，逐条回退: %s", e)
            cancelled = _export_all(None)
    else:
        cancelled = _export_all(None)

    if cancelled:
        logger.info("Cancelled by user after %d PDFs", len(generated))
        return {
            "batch_ok": True,
            "message": f"已中断（已生成 {len(generated)} 个 PDF）",
            "report": str(report_path),
            "generated": generated,
            "container_count": len(generated),
            "bill_name": batch.bill_name,
            "system_total": str(batch.system_grand_total),
            "summary_total": str(batch.summary_grand_total),
            "render_mode": mode,
            "cancelled": True,
        }

    return {
        "batch_ok": True,
        "message": batch.batch_message,
        "report": str(report_path),
        "generated": generated,
        "container_count": len(generated),
        "bill_name": batch.bill_name,
        "system_total": str(batch.system_grand_total),
        "summary_total": str(batch.summary_grand_total),
        "render_mode": mode,
        "cancelled": False,
    }
