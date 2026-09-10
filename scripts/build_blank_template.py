# -*- coding: utf-8 -*-
"""从 samples 已确认账单提取空白发票 Excel 模板。"""
from __future__ import annotations

from copy import copy
from pathlib import Path

import openpyxl
from openpyxl.cell.cell import MergedCell
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parents[1]
SAMPLES = ROOT / "samples"
OUT = ROOT / "templates" / "invoice_blank.xlsx"


def copy_cell(src_cell, dst_cell):
    dst_cell.value = src_cell.value
    if src_cell.has_style:
        dst_cell.font = copy(src_cell.font)
        dst_cell.border = copy(src_cell.border)
        dst_cell.fill = copy(src_cell.fill)
        dst_cell.number_format = src_cell.number_format
        dst_cell.protection = copy(src_cell.protection)
        dst_cell.alignment = copy(src_cell.alignment)


def main():
    xlsx = next(SAMPLES.glob("*.xlsx"))
    wb = openpyxl.load_workbook(xlsx)
    ws = wb[wb.sheetnames[0]]

    out_wb = openpyxl.Workbook()
    out = out_wb.active
    out.title = "INVOICE"

    for c in range(1, 16):
        letter = get_column_letter(c)
        if letter in ws.column_dimensions and ws.column_dimensions[letter].width:
            out.column_dimensions[letter].width = ws.column_dimensions[letter].width

    for mr in ws.merged_cells.ranges:
        if mr.max_row <= 14:
            out.merge_cells(str(mr))

    for r in range(1, 15):
        if ws.row_dimensions[r].height:
            out.row_dimensions[r].height = ws.row_dimensions[r].height
        for c in range(1, 16):
            sc = ws.cell(r, c)
            if isinstance(sc, MergedCell):
                continue
            copy_cell(sc, out.cell(r, c))

    # 动态字段留空
    out["H2"] = ""
    out["H5"] = ""

    for r in range(15, 41):
        for c in range(1, 9):
            sc = ws.cell(15, c)
            dc = out.cell(r, c)
            if sc.has_style:
                dc.font = copy(sc.font)
                dc.border = copy(sc.border)
                dc.fill = copy(sc.fill)
                dc.number_format = sc.number_format
                dc.alignment = copy(sc.alignment)
            dc.value = None

    out["G42"] = "Total Amount:"
    out["H42"] = 0
    out["H42"].number_format = '"$"#,##0.00'
    out["A44"] = "If you have any questions about this invoice, please contact"
    out["A45"] = "[ ERIC, 626-669-5398 ]"

    out.page_setup.orientation = "landscape"
    out.page_setup.paperSize = out.PAPERSIZE_A4
    out.print_area = "A1:H45"

    OUT.parent.mkdir(parents=True, exist_ok=True)
    out_wb.save(OUT)
    print("saved", OUT)


if __name__ == "__main__":
    main()
