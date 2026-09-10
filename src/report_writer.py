from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

from .models import BatchResult


def write_report(batch: BatchResult, report_path: Path) -> Path:
    report_path = Path(report_path)
    report_path.parent.mkdir(parents=True, exist_ok=True)

    wb = Workbook()
    ws = wb.active
    ws.title = "校验报告"

    ws.append(
        [
            "柜号",
            "明细行数",
            "系统合计",
            "汇总合计",
            "差额",
            "状态",
            "说明",
            "Invoice No",
            "输出文件",
        ]
    )
    for cell in ws[1]:
        cell.font = Font(bold=True)

    red = PatternFill("solid", fgColor="FFCCCC")
    yellow = PatternFill("solid", fgColor="FFF2CC")
    green = PatternFill("solid", fgColor="C6EFCE")

    inv_map = {i.container_no: i for i in batch.invoices}
    for v in batch.validations:
        inv = inv_map.get(v.container_no)
        ws.append(
            [
                v.container_no,
                len(inv.lines) if inv else 0,
                float(v.system_total),
                float(v.summary_total) if v.summary_total is not None else None,
                float(v.diff) if v.diff is not None else None,
                v.status,
                v.message,
                inv.invoice_no if inv else "",
                f"{v.container_no}.pdf" if v.status != "FAIL" and batch.batch_ok else "",
            ]
        )
        fill = {"OK": green, "WARN": yellow, "FAIL": red}.get(v.status)
        if fill:
            for cell in ws[ws.max_row]:
                cell.fill = fill

    ws.append([])
    ws.append(["账单名称", batch.bill_name])
    ws.append(["账单日期", batch.invoice_date])
    ws.append(["系统批次合计", float(batch.system_grand_total)])
    ws.append(
        [
            "汇总批次合计",
            float(batch.summary_grand_total) if batch.summary_grand_total is not None else None,
        ]
    )
    ws.append(["批次是否放行", "YES" if batch.batch_ok else "NO"])
    ws.append(["批次说明", batch.batch_message])

    wq = wb.create_sheet("quarantine")
    wq.append(["行号", "原因", "详情"])
    for cell in wq[1]:
        cell.font = Font(bold=True)
    for q in batch.quarantine:
        wq.append([q.row_number, q.reason, str(q.raw)])

    schema = wb.create_sheet("schema_warnings")
    schema.append(["行号", "类型", "详情"])
    schema.append([0, "主表", batch.main_sheet_name if hasattr(batch, "main_sheet_name") else ""])
    for cell in schema[1]:
        cell.font = Font(bold=True)
    for q in batch.schema_warnings:
        schema.append([q.row_number, q.reason, str(q.raw)])

    for sheet in (ws, wq, schema):
        for col in sheet.columns:
            sheet.column_dimensions[col[0].column_letter].width = 18

    wb.save(report_path)
    return report_path
