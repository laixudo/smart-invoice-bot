from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date
from decimal import Decimal
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

from .money import money_diff, to_decimal
from .models import ContainerInvoice, LineItem, QuarantineRow


HEADER_ALIASES = {
    "appointment id": "appointment_id",
    "total volume": "total_volume",
    "cbm": "cbm",
    "destination": "destination",
    "isa": "isa",
    "qty": "qty",
    "rate": "rate",
    "amount": "amount",
    "最终结算金额": "final_settlement",
    "测试反馈": "reconciliation_feedback",
    "备注": "note",
    "九方对账反馈": "reconciliation_feedback",
    "ar发票": "ar_invoice",
    "ar发票号": "ar_invoice",
    "合计金额": "container_total",
    "invoice编号": "legacy_invoice_no",
}


@dataclass
class ParseOutput:
    invoices: list[ContainerInvoice]
    quarantine: list[QuarantineRow]
    bill_name: str
    invoice_date: str
    summary_totals: dict[str, Decimal]
    summary_grand_total: Decimal | None
    main_sheet_name: str
    schema_warnings: list[QuarantineRow]


def _cell_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def find_header_row(ws: Worksheet, scan_rows: int, required: list[str]) -> tuple[int, dict[str, int]]:
    max_scan = min(scan_rows, ws.max_row or 1)
    for r in range(1, max_scan + 1):
        col_map: dict[str, int] = {}
        texts: list[str] = []
        for c in range(1, min((ws.max_column or 1) + 1, 40)):
            raw = ws.cell(r, c).value
            if raw is None:
                continue
            t = _cell_text(raw)
            texts.append(t.lower())
            key = HEADER_ALIASES.get(t.lower())
            if key:
                col_map[key] = c
        joined = " | ".join(texts)
        if all(k.lower() in joined for k in required):
            if {"appointment_id", "amount", "destination"} <= set(col_map):
                return r, col_map
    raise ValueError(
        f"未找到表头（已扫描前 {scan_rows} 行）。需要同时包含: {', '.join(required)}"
    )


def _build_merge_map(ws: Worksheet, col: int) -> dict[int, Any]:
    out: dict[int, Any] = {}
    for mr in ws.merged_cells.ranges:
        if mr.min_col == mr.max_col == col:
            value = ws.cell(mr.min_row, col).value
            for r in range(mr.min_row, mr.max_row + 1):
                out[r] = value
    return out


def _build_vertical_merge_spans(ws: Worksheet, col: int) -> dict[int, tuple[int, int, Any]]:
    """Map rows in vertically merged cells to (first row, last row, anchor value)."""
    out: dict[int, tuple[int, int, Any]] = {}
    for mr in ws.merged_cells.ranges:
        if mr.min_col == mr.max_col == col and mr.max_row > mr.min_row:
            value = ws.cell(mr.min_row, col).value
            for r in range(mr.min_row, mr.max_row + 1):
                out[r] = (mr.min_row, mr.max_row, value)
    return out


def _legacy_container_id(value: Any) -> str:
    """Extract the real 11-character container number from legacy mixed text."""
    match = re.search(r"(?<![A-Z0-9])([A-Z]{4}\d{7})(?!\d)", _cell_text(value).upper())
    return match.group(1) if match else ""


def _legacy_total_volume(container_no: str, *values: Any) -> Decimal | None:
    """Read volume appended after a container number, ignoring 20/40/45 foot markers."""
    if not container_no:
        return None
    for value in values:
        text = _cell_text(value).upper()
        container_at = text.find(container_no)
        if container_at < 0:
            continue
        tail = text[container_at + len(container_no):]
        tail = re.sub(r"-?\s*(?:20|40|45)\s*(?:HC|HQ|GP)\b", " ", tail)
        numbers = [to_decimal(v) for v in re.findall(r"\d+(?:\.\d+)?", tail)]
        candidates = [v for v in numbers if v is not None and v not in {Decimal("20"), Decimal("40"), Decimal("45")}]
        if candidates:
            return candidates[-1]
    return None


def _legacy_container_start_overrides(
    ws: Worksheet,
    header_row: int,
    col_map: dict[str, int],
) -> dict[int, str]:
    """Move a legacy group start upward when charge cells merge across its A-cell boundary."""
    relevant_cols = {col_map[name] for name in ("qty", "rate", "amount") if name in col_map}
    overrides: dict[int, str] = {}
    for r in range(header_row + 1, (ws.max_row or header_row) + 1):
        container_no = _legacy_container_id(ws.cell(r, col_map["appointment_id"]).value)
        if not container_no:
            continue
        start_row = r
        for merged in ws.merged_cells.ranges:
            if (
                merged.min_col == merged.max_col
                and merged.min_col in relevant_cols
                and merged.min_row < r <= merged.max_row
            ):
                start_row = min(start_row, merged.min_row)
        if start_row < r:
            overrides[start_row] = container_no
    return overrides


_LEGACY_CBM_SUFFIX = re.compile(
    r"^(?P<destination>.+?)\s*-\s*(?:GY\s*)?(?P<cbm>\d+(?:\.\d+)?)\s*(?:CBM)?$",
    re.I,
)


def _legacy_destination_cbm_isa(
    destination_raw: Any,
    isa_raw: Any,
    container_no: str,
) -> tuple[str, Decimal | None, str]:
    """Split legacy DESTINATION text such as 'FTW1 - 6.98' into fields."""
    destination_text = _cell_text(destination_raw)
    isa_text = _cell_text(isa_raw)

    # Some first/fee rows put container metadata in DESTINATION and the real
    # charge description in ISA (for example: container / 超重费 42716LBS).
    if container_no and container_no in destination_text.upper():
        source = isa_text
        output_isa = ""
    else:
        source = destination_text
        output_isa = _format_isa(isa_raw)

    match = _LEGACY_CBM_SUFFIX.match(source)
    if not match:
        return source, None, output_isa
    return (
        match.group("destination").strip(),
        to_decimal(match.group("cbm")),
        output_isa,
    )


def _pick_main_sheet(wb) -> Worksheet:
    for name in wb.sheetnames:
        upper = name.upper()
        if name.startswith("#") or "SMART" in upper or "DS-" in upper:
            return wb[name]
    return wb[wb.sheetnames[0]]


def _extract_bill_name(ws: Worksheet, file_path: Path) -> str:
    for r, c in ((2, 8), (1, 8), (2, 1)):
        v = _cell_text(ws.cell(r, c).value)
        if v.startswith("#DS-") or "SMART" in v.upper():
            return v
    m = re.search(r"(#DS-\d+-SMART)", file_path.stem, re.I)
    return m.group(1) if m else file_path.stem


def _date_from_bill_name(bill_name: str) -> str | None:
    m = re.search(r"#DS-(\d{2})(\d{2})(\d{4})", bill_name)
    if not m:
        return None
    return f"{m.group(1)}/{m.group(2)}/{m.group(3)}"


def _date_from_excel(ws: Worksheet) -> str | None:
    for r in range(1, 12):
        for c in range(1, 12):
            if _cell_text(ws.cell(r, c).value).upper() != "DATE":
                continue
            v = ws.cell(r + 1, c).value
            if v is None:
                continue
            if hasattr(v, "strftime"):
                return v.strftime("%m/%d/%Y")
            text = _cell_text(v)
            m = re.match(r"^(\d{2})/(\d{2})(\d{4})$", text)
            if m:
                return f"{m.group(1)}/{m.group(2)}/{m.group(3)}"
            if re.match(r"^\d{2}/\d{2}/\d{4}$", text):
                return text
            return text
    return None


def resolve_invoice_date(cfg: dict, ws: Worksheet, bill_name: str) -> str:
    source = cfg.get("date_source", "bill_name")
    if source == "excel_header":
        return _date_from_excel(ws) or _date_from_bill_name(bill_name) or date.today().strftime("%m/%d/%Y")
    if source == "today":
        return date.today().strftime("%m/%d/%Y")
    return _date_from_bill_name(bill_name) or _date_from_excel(ws) or date.today().strftime("%m/%d/%Y")


def _contains_keyword(text: str, keywords: list[str]) -> bool:
    low = text.lower()
    return any(kw.lower() in low for kw in keywords)


def _configured_fee_amount(destination: str, note: str, feedback: str, qty: Decimal | None, amount: Decimal | None, cfg: dict) -> Decimal | None:
    rules = cfg.get("fee_rules", {})
    chassis = rules.get("chassis", {})
    if chassis.get("enabled", False) and _contains_keyword(destination, chassis.get("destination_keywords", [])):
        match = re.search(chassis.get("days_regex", r"(\d+)"), note)
        if not match:
            return amount
        days = Decimal(match.group(1))
        free_days = to_decimal(chassis.get("free_days", 5)) or Decimal("5")
        daily_rate = to_decimal(chassis.get("daily_rate", 40)) or Decimal("0")
        return max(Decimal("0"), days - free_days) * daily_rate
    pallet = rules.get("pallet", {})
    if pallet.get("enabled", False) and _contains_keyword(feedback, pallet.get("feedback_keywords", [])):
        unit_rate = to_decimal(pallet.get("unit_rate", 100)) or Decimal("0")
        return (qty or Decimal("0")) * unit_rate
    return amount


def _format_isa(isa_raw: Any) -> str:
    if isa_raw is None:
        return ""
    if isinstance(isa_raw, int):
        return str(isa_raw)
    if isinstance(isa_raw, float):
        if isa_raw.is_integer():
            return str(int(isa_raw))
        return format(isa_raw, "f").rstrip("0").rstrip(".")
    return str(isa_raw).strip()


def _read_summary(wb, container_re: re.Pattern) -> tuple[dict[str, Decimal], Decimal | None]:
    if "汇总" not in wb.sheetnames:
        return {}, None
    ws = wb["汇总"]
    totals: dict[str, Decimal] = {}
    grand: Decimal | None = None
    for r in range(1, (ws.max_row or 1) + 1):
        aid = _cell_text(ws.cell(r, 1).value)
        amt = to_decimal(ws.cell(r, 2).value)
        if aid and container_re.match(aid) and amt is not None:
            totals[aid] = amt
        elif not aid and amt is not None and r > 1:
            grand = amt
    if grand is None and totals:
        grand = sum(totals.values(), Decimal("0"))
    return totals, grand


def _make_invoice_no(cfg: dict, ar: str | None, seq: int, invoice_date: str) -> str:
    if cfg.get("invoice_no_source", "ar_column") == "ar_column" and ar:
        return ar
    parts = invoice_date.split("/")
    yyyymm = f"{parts[2]}{parts[0]}" if len(parts) == 3 else date.today().strftime("%Y%m")
    return f"{yyyymm}{seq:02d}"


def _schema_warnings(ws: Worksheet, header_row: int, cfg: dict) -> list[QuarantineRow]:
    rules = cfg.get("schema_detection", {})
    if not rules.get("enabled", True):
        return []
    allowed = {str(v).strip().lower() for v in rules.get("allowed_headers", [])}
    allowed.update(HEADER_ALIASES)
    warnings: list[QuarantineRow] = []
    for col in range(1, (ws.max_column or 1) + 1):
        value = _cell_text(ws.cell(header_row, col).value)
        if value and value.lower() not in allowed:
            warnings.append(QuarantineRow(header_row, "未知表头", {"column": ws.cell(header_row, col).coordinate, "value": value}))
    for col, dim in ws.column_dimensions.items():
        if dim.hidden:
            warnings.append(QuarantineRow(header_row, "检测到隐藏列", {"column": col, "header": _cell_text(ws[f"{col}{header_row}"].value)}))
    return warnings


def parse_excel(path: str | Path, cfg: dict) -> ParseOutput:
    file_path = Path(path)
    wb = load_workbook(file_path, data_only=True)
    ws = _pick_main_sheet(wb)
    header_row, col_map = find_header_row(
        ws,
        int(cfg.get("header_scan_rows", 20)),
        list(cfg.get("header_required_keywords", ["APPOINTMENT ID", "DESTINATION", "AMOUNT"])),
    )
    # 核心模板字段固定：少任何一个都不能安全出账。
    core_columns = list(cfg.get("core_columns", [
        "appointment_id", "total_volume", "cbm", "destination", "isa", "qty", "rate", "amount",
    ]))
    legacy_layout = (
        "container_total" in col_map
        and "legacy_invoice_no" in col_map
        and "total_volume" not in col_map
        and "cbm" not in col_map
    )
    missing_core = [name for name in core_columns if name not in col_map]
    if legacy_layout:
        missing_core = [name for name in missing_core if name not in {"total_volume", "cbm"}]
    if missing_core:
        raise ValueError(f"核心模板字段缺失: {', '.join(missing_core)}")

    container_re = re.compile(cfg.get("container_id_regex", r"^[A-Z]{4}\d{7}(?:-\d+)?$"))
    surcharge_kw = list(cfg.get("surcharge_keywords", []))
    tail_kw = list(cfg.get("tail_block_keywords", []))
    unclassified = cfg.get("unclassified_row_policy", "inherit")

    merge_id = _build_merge_map(ws, col_map["appointment_id"])
    vol_col = col_map.get("total_volume")
    merge_vol = _build_merge_map(ws, vol_col) if vol_col else {}
    amount_merge_spans = (
        _build_vertical_merge_spans(ws, col_map["amount"])
        if legacy_layout
        else {}
    )
    legacy_start_overrides = (
        _legacy_container_start_overrides(ws, header_row, col_map)
        if legacy_layout
        else {}
    )

    bill_name = _extract_bill_name(ws, file_path)
    invoice_date = resolve_invoice_date(cfg, ws, bill_name)
    bill_to = cfg.get("bill_to", "SMART")

    quarantine: list[QuarantineRow] = []
    schema_warnings = _schema_warnings(ws, header_row, cfg)
    if legacy_layout:
        schema_warnings.append(
            QuarantineRow(
                header_row,
                "已启用旧版模板兼容解析",
                {
                    "total_volume": "从柜号行混合文本提取，缺失时按明细 CBM 合计",
                    "cbm": "从 DESTINATION/ISA 文本末尾提取",
                    "invoice_no_column": ws.cell(header_row, col_map["legacy_invoice_no"]).coordinate,
                },
            )
        )
    invoices: list[ContainerInvoice] = []
    current: ContainerInvoice | None = None
    seq = 0

    # “最终结算金额”是辅助审计列，不参与正式出账金额计算。客户文件有时会
    # 隐藏该列并复用它记录对账反馈；这种情况下不要把反馈文字当成金额阻断整批。
    # 核心账单列仍按原逻辑读取，不受隐藏状态影响。
    final_settlement_col = col_map.get("final_settlement")
    final_settlement_col_hidden = bool(
        final_settlement_col
        and ws.column_dimensions[ws.cell(header_row, final_settlement_col).column_letter].hidden
    )
    invoice_no_col = (
        col_map.get("legacy_invoice_no") if legacy_layout else col_map.get("ar_invoice")
    )
    container_total_col = col_map.get("container_total") if legacy_layout else None

    max_row = ws.max_row or header_row
    for r in range(header_row + 1, max_row + 1):
        raw_id = ws.cell(r, col_map["appointment_id"]).value
        merged_id = merge_id.get(r)
        displayed_id = (
            (legacy_start_overrides.get(r) or _legacy_container_id(raw_id))
            if legacy_layout
            else _cell_text(raw_id)
        )
        effective_id = displayed_id or (
            _legacy_container_id(merged_id) if legacy_layout else _cell_text(merged_id)
        )

        destination_raw = ws.cell(r, col_map["destination"]).value
        isa_raw = ws.cell(r, col_map["isa"]).value if "isa" in col_map else None
        if legacy_layout:
            dest, legacy_cbm, isa = _legacy_destination_cbm_isa(
                destination_raw,
                isa_raw,
                effective_id,
            )
        else:
            dest = _cell_text(destination_raw)
            legacy_cbm = None
            isa = _format_isa(isa_raw)
        qty = to_decimal(ws.cell(r, col_map["qty"]).value) if "qty" in col_map else None
        rate = to_decimal(ws.cell(r, col_map["rate"]).value) if "rate" in col_map else None
        amount_raw = ws.cell(r, col_map["amount"]).value if "amount" in col_map else None
        amount = to_decimal(amount_raw)
        if legacy_layout and r in amount_merge_spans:
            merge_start, _, merge_value = amount_merge_spans[r]
            # A vertically merged AMOUNT is a source group total, not a value
            # that should be repeated or recomputed for every destination row.
            amount = to_decimal(merge_value) if r == merge_start else Decimal("0")
        cbm = (
            legacy_cbm
            if legacy_layout
            else (to_decimal(ws.cell(r, col_map["cbm"]).value) if "cbm" in col_map else None)
        )
        note = _cell_text(ws.cell(r, col_map["note"]).value) if "note" in col_map else ""
        feedback = _cell_text(ws.cell(r, col_map["reconciliation_feedback"]).value) if "reconciliation_feedback" in col_map else ""
        legacy_settlement_raw = (
            ws.cell(r, final_settlement_col).value
            if legacy_layout and final_settlement_col
            else None
        )
        if legacy_layout and legacy_settlement_raw not in (None, "") and to_decimal(legacy_settlement_raw) is None:
            feedback = _cell_text(legacy_settlement_raw)
        expected_fee_amount = _configured_fee_amount(dest, note, feedback, qty, amount, cfg)
        # 源表金额是唯一可出账金额；费用规则只做审计，不得擅自改写客户账单。
        if (
            expected_fee_amount is not None
            and amount is not None
            and expected_fee_amount != amount
        ):
            schema_warnings.append(
                QuarantineRow(
                    r,
                    "费用规则与源表金额不一致",
                    {
                        "destination": dest,
                        "source_amount": str(amount),
                        "rule_amount": str(expected_fee_amount),
                        "note": note,
                        "feedback": feedback,
                    },
                )
            )

        if legacy_layout:
            volume = _legacy_total_volume(effective_id, raw_id, merged_id, destination_raw)
        else:
            vol_raw = ws.cell(r, vol_col).value if vol_col else None
            if vol_raw is None and vol_col:
                vol_raw = merge_vol.get(r)
            volume = to_decimal(vol_raw)

        final_settlement = None
        final_settlement_row = None
        if legacy_layout:
            final_settlement = to_decimal(legacy_settlement_raw)
            if final_settlement is None and container_total_col:
                final_settlement = to_decimal(ws.cell(r, container_total_col).value)
            if final_settlement is not None:
                final_settlement_row = r
        elif final_settlement_col and not final_settlement_col_hidden:
            raw_final_settlement = ws.cell(r, final_settlement_col).value
            if raw_final_settlement not in (None, ""):
                final_settlement = to_decimal(raw_final_settlement)
                final_settlement_row = r
                if final_settlement is None:
                    schema_warnings.append(
                        QuarantineRow(
                            r,
                            "最终结算金额非数字",
                            {"value": _cell_text(raw_final_settlement)},
                        )
                    )

        ar = None
        if invoice_no_col:
            ar_v = ws.cell(r, invoice_no_col).value
            if ar_v is not None:
                ar = str(ar_v).strip()
                if ar.endswith(".0"):
                    ar = ar[:-2]

        row_blob = " ".join(_cell_text(ws.cell(r, c).value) for c in range(1, 9))
        if _contains_keyword(row_blob, tail_kw):
            if "total amount" in row_blob.lower() or _contains_keyword(row_blob, ["总计", "合计"]):
                break
            quarantine.append(QuarantineRow(r, "尾部/全局关键词", {"blob": row_blob[:200]}))
            break

        if all(
            x is None or x == ""
            for x in (raw_id, dest, amount, cbm, qty, rate)
        ):
            continue

        if displayed_id and not container_re.match(displayed_id):
            quarantine.append(
                QuarantineRow(
                    r,
                    f"非法柜号格式: {displayed_id}",
                    {"destination": dest, "amount": str(amount)},
                )
            )
            continue

        if effective_id and container_re.match(effective_id):
            if current is None or current.container_no != effective_id:
                if current and current.lines:
                    invoices.append(current)
                seq += 1
                current = ContainerInvoice(
                    container_no=effective_id,
                    total_volume=volume,
                    invoice_no=_make_invoice_no(cfg, ar, seq, invoice_date),
                    bill_name=bill_name,
                    invoice_date=invoice_date,
                    bill_to=bill_to,
                    source_final_settlement=final_settlement,
                    source_final_settlement_row=final_settlement_row,
                    source_feedback=feedback,
                    source_feedback_row=r if feedback else None,
                )
            else:
                if current.total_volume is None and volume is not None:
                    current.total_volume = volume
                if ar and (
                    legacy_layout
                    or not current.invoice_no
                    or current.invoice_no.endswith("00")
                ):
                    current.invoice_no = ar
        else:
            has_content = amount is not None or bool(dest)
            if not has_content:
                continue
            if current is None:
                quarantine.append(
                    QuarantineRow(r, "无当前柜号可归属", {"destination": dest, "amount": str(amount)})
                )
                continue
            is_surcharge = _contains_keyword(dest, surcharge_kw)
            if not (is_surcharge or unclassified == "inherit"):
                if unclassified == "fail":
                    raise ValueError(f"第 {r} 行无法归属柜号: {dest}")
                quarantine.append(
                    QuarantineRow(r, "未分类空柜号行", {"destination": dest, "amount": str(amount)})
                )
                continue

        if current is None:
            continue
        if amount is None:
            if qty is not None and rate is not None:
                amount = qty * rate
            else:
                quarantine.append(
                    QuarantineRow(r, "AMOUNT 为空且无法用 QTY×RATE 计算", {"destination": dest})
                )
                continue

        current.lines.append(
            LineItem(
                cbm=cbm,
                destination=dest or "",
                isa=isa,
                qty=qty,
                rate=rate,
                amount=amount,
                source_row=r,
            )
        )
        if current.total_volume is None and volume is not None:
            current.total_volume = volume

    if current and current.lines:
        invoices.append(current)

    if legacy_layout:
        for invoice in invoices:
            if invoice.total_volume is not None:
                continue
            cbm_values = [line.cbm for line in invoice.lines if line.cbm is not None]
            if cbm_values:
                invoice.total_volume = sum(cbm_values, Decimal("0"))
                schema_warnings.append(
                    QuarantineRow(
                        invoice.lines[0].source_row,
                        "旧版模板总体积由明细CBM合计",
                        {
                            "container_no": invoice.container_no,
                            "total_volume": str(invoice.total_volume),
                        },
                    )
                )
            else:
                schema_warnings.append(
                    QuarantineRow(
                        invoice.lines[0].source_row,
                        "旧版模板无法确定总体积",
                        {"container_no": invoice.container_no},
                    )
                )

    # The customer-approved rule is: add raw source amounts within each container,
    # then round once. Keep source settlement cells as an auditable reference.
    for invoice in invoices:
        if invoice.source_final_settlement is not None:
            diff = money_diff(invoice.total_amount, invoice.source_final_settlement)
            if diff:
                schema_warnings.append(
                    QuarantineRow(
                        invoice.source_final_settlement_row or 0,
                        "最终结算金额与正式柜合计不一致",
                        {
                            "container_no": invoice.container_no,
                            "source_settlement": str(invoice.source_final_settlement),
                            "official_container_total": str(invoice.total_amount),
                            "difference": str(diff),
                        },
                    )
                )
        # Historical reconciliation comments are checked against the customer-approved
        # per-container total, so a genuine conflict remains visible in the report.
        feedback_match = re.search(
            r"(?:应为|应该是|应是)\s*\$?\s*(-?[\d,]+(?:\.\d+)?)",
            invoice.source_feedback,
        )
        if feedback_match:
            feedback_amount = to_decimal(feedback_match.group(1))
            if feedback_amount is not None:
                diff = money_diff(invoice.total_amount, feedback_amount)
                if diff:
                    schema_warnings.append(
                        QuarantineRow(
                            invoice.source_feedback_row or 0,
                            "测试反馈金额与正式柜合计不一致",
                            {
                                "container_no": invoice.container_no,
                                "feedback_amount": str(feedback_amount),
                                "official_container_total": str(invoice.total_amount),
                                "difference": str(diff),
                                "policy": "柜内原始明细相加后四舍五入",
                            },
                        )
                    )

    # 隐藏汇总表通常为人工辅助数据；默认不作为出账与校验依据。
    # 需要时可在配置中显式启用，避免手工汇总值阻断明细账单生成。
    if cfg.get("summary_validation", {}).get("enabled", False):
        summary_totals, summary_grand = _read_summary(wb, container_re)
    else:
        summary_totals, summary_grand = {}, None
    return ParseOutput(
        invoices=invoices,
        quarantine=quarantine,
        bill_name=bill_name,
        invoice_date=invoice_date,
        summary_totals=summary_totals,
        summary_grand_total=summary_grand,
        main_sheet_name=ws.title,
        schema_warnings=schema_warnings,
    )
