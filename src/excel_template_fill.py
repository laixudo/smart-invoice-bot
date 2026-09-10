"""把单柜数据填入 Excel 发票模板（真正的『填模板』）。"""

from __future__ import annotations

from copy import copy
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from .models import ContainerInvoice
from .money import quantize

from .paths import resource_root

DEFAULT_TEMPLATE = resource_root() / "templates" / "invoice_blank.xlsx"


def fill_invoice_xlsx(
    invoice: ContainerInvoice,
    out_xlsx: Path,
    template_path: Path | None = None,
) -> Path:
    tpl = Path(template_path) if template_path else DEFAULT_TEMPLATE
    if not tpl.exists():
        raise FileNotFoundError(f"发票模板不存在: {tpl}")

    wb = load_workbook(tpl)
    ws = wb.active

    # 抬头动态字段
    ws["H2"] = invoice.bill_name
    ws["H5"] = invoice.invoice_date
    ws["A10"] = invoice.bill_to

    # Invoice no —— 放在表头右侧（与样例一致：G12/H12 一带）
    # 模板若无该单元格文案，直接写入
    ws["G12"] = "Invoice no："
    ws["H12"] = invoice.invoice_no

    thin = Side(style="thin", color="808080")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="center")

    start_row = 15
    max_slot = 40  # 模板预留行
    # 清空旧数据区
    for r in range(start_row, max_slot + 1):
        for c in range(1, 9):
            ws.cell(r, c).value = None

    vol = float(invoice.total_volume) if invoice.total_volume is not None else None
    lines = invoice.display_lines
    for i, ln in enumerate(lines):
        r = start_row + i
        if r > max_slot:
            raise ValueError(
                f"{invoice.container_no} 明细 {len(invoice.lines)} 行超过模板容量 {max_slot - start_row + 1}"
            )
        ws.cell(r, 1).value = invoice.container_no if i == 0 else None
        ws.cell(r, 2).value = vol
        ws.cell(r, 3).value = float(ln.cbm) if ln.cbm is not None else None
        ws.cell(r, 4).value = ln.destination
        ws.cell(r, 5).value = ln.isa
        ws.cell(r, 6).value = float(ln.qty) if ln.qty is not None else None
        ws.cell(r, 7).value = float(ln.rate) if ln.rate is not None else None
        ws.cell(r, 8).value = float(quantize(ln.amount))
        for c in range(1, 9):
            cell = ws.cell(r, c)
            cell.border = border
            cell.alignment = left if c == 1 else center
            if c == 8:
                cell.number_format = "0.00"
            elif c == 7:
                cell.number_format = "0"
            elif c == 6:
                cell.number_format = "0.00000000"
            elif c in (2, 3):
                cell.number_format = "0.##"

    # 合并柜号列（视觉接近黄金样例）
    end_row = start_row + len(lines) - 1
    if len(lines) > 1:
        # 先拆可能冲突的合并
        to_unmerge = [
            str(m)
            for m in ws.merged_cells.ranges
            if m.min_col in (1, 2) and m.min_row >= start_row
        ]
        for m in to_unmerge:
            ws.unmerge_cells(m)
        ws.merge_cells(start_row=start_row, start_column=1, end_row=end_row, end_column=1)
        ws.merge_cells(start_row=start_row, start_column=2, end_row=end_row, end_column=2)
        ws.cell(start_row, 1).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(start_row, 2).alignment = Alignment(horizontal="center", vertical="center")

    total_row = end_row + 2
    ws.cell(total_row, 7).value = "Total Amount:"
    ws.cell(total_row, 7).font = Font(bold=True, color="FFFFFF")
    ws.cell(total_row, 7).fill = PatternFill("solid", fgColor="808080")
    ws.cell(total_row, 8).value = float(quantize(invoice.total_amount))
    ws.cell(total_row, 8).number_format = '"$"#,##0.00'
    ws.cell(total_row, 8).font = Font(bold=True, color="FFFFFF")
    ws.cell(total_row, 8).fill = PatternFill("solid", fgColor="808080")

    footer1 = total_row + 2
    ws.cell(footer1, 1).value = "If you have any questions about this invoice, please contact"
    ws.cell(footer1 + 1, 1).value = "[ ERIC, 626-669-5398 ]"

    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    try:
        ws.sheet_properties.pageSetUpPr.fitToPage = True
    except Exception:
        pass
    ws.print_area = f"A1:H{footer1 + 1}"

    out_xlsx = Path(out_xlsx)
    out_xlsx.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_xlsx)
    return out_xlsx
