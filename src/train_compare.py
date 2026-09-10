"""用本地黄金样例校准：字段对账 + 版式基础检查（非 ML 训练）。"""

from __future__ import annotations

import re
import sys
from decimal import Decimal
from pathlib import Path

import pdfplumber

from .config_loader import load_config
from .excel_parser import parse_excel
from .money import format_money, quantize
from .pdf_renderer import PAGE_SIZE, render_invoice_pdf
from .validator import validate_batch

ROOT = Path(__file__).resolve().parents[1]
SAMPLES = ROOT / "samples"
GOLD_PDF = SAMPLES / "TIIU6724883.gold.pdf"


def _find_sample_xlsx() -> Path:
    for p in SAMPLES.glob("*.xlsx"):
        if "SMART" in p.name.upper() or p.name.startswith("#DS"):
            return p
    raise FileNotFoundError(f"samples/ 下未找到已确认 Excel: {SAMPLES}")


def _extract_fields(pdf_path: Path) -> dict:
    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[0]
        text = page.extract_text() or ""
        size = (round(page.width, 2), round(page.height, 2))
    bill = re.search(r"(#DS-\d+-SMART)", text)
    inv = re.search(r"Invoice no[:：]\s*(\d+)", text, re.I)
    total = re.search(r"Total Amount:\s*(\$?[\d,]+\.\d{2})", text, re.I)
    date = re.search(r"(\d{2}/\d{2}/\d{4})", text)
    container = re.search(r"\b([A-Z]{4}\d{7})\b", text)
    return {
        "pages": len(pdfplumber.open(pdf_path).pages),
        "size": size,
        "landscape": size[0] > size[1],
        "bill_name": bill.group(1) if bill else None,
        "invoice_no": inv.group(1) if inv else None,
        "total": total.group(1).replace("$", "") if total else None,
        "date": date.group(1) if date else None,
        "container": container.group(1) if container else None,
        "text": text,
    }


def run_train() -> int:
    xlsx = _find_sample_xlsx()
    if not GOLD_PDF.exists():
        print(f"缺少黄金 PDF: {GOLD_PDF}", file=sys.stderr)
        return 2

    cfg = load_config()
    parsed = parse_excel(xlsx, cfg)
    batch = validate_batch(parsed, cfg)
    by_id = {i.container_no: i for i in parsed.invoices}
    ti = by_id.get("TIIU6724883")
    if ti is None:
        print("解析结果中无 TIIU6724883", file=sys.stderr)
        return 1

    out_dir = ROOT / "output" / "_train"
    out_dir.mkdir(parents=True, exist_ok=True)
    gen_pdf = out_dir / "TIIU6724883.pdf"
    render_invoice_pdf(ti, gen_pdf, cfg)

    gold = _extract_fields(GOLD_PDF)
    gen = _extract_fields(gen_pdf)

    checks = [
        ("柜号", gold["container"], gen["container"]),
        ("账单名", gold["bill_name"], gen["bill_name"]),
        ("Invoice no", gold["invoice_no"], gen["invoice_no"]),
        ("日期", gold["date"], gen["date"]),
        ("Total", gold["total"].replace(",", "") if gold["total"] else None,
         gen["total"].replace(",", "") if gen["total"] else None),
        ("横向版式", True, gen["landscape"]),
        ("页宽≈A4横", True, abs(gen["size"][0] - PAGE_SIZE[0]) < 2),
        ("页高≈A4横", True, abs(gen["size"][1] - PAGE_SIZE[1]) < 2),
        ("Excel合计", "4862.71", str(quantize(ti.total_amount))),
        ("批次放行", True, batch.batch_ok),
    ]

    print("==== 黄金样例校准 ====")
    print(f"Excel: {xlsx.name}")
    print(f"Gold : {GOLD_PDF.name} size={gold['size']} pages={gold['pages']}")
    print(f"Gen  : {gen_pdf.name} size={gen['size']} pages={gen['pages']}")
    print(f"期望 Total: {format_money(Decimal('4862.71'))}")
    ok = True
    for name, expect, actual in checks:
        passed = expect == actual
        if not passed:
            ok = False
        mark = "PASS" if passed else "FAIL"
        print(f"[{mark}] {name}: expect={expect!r} actual={actual!r}")

    # 明细关键词
    for kw in ["PHX5", "LFT1", "DEN8", "FAT2", "CMH2", "SMART"]:
        if kw not in gen["text"]:
            ok = False
            print(f"[FAIL] 生成 PDF 缺少关键词: {kw}")
        else:
            print(f"[PASS] 含关键词: {kw}")

    print("====", "全部通过" if ok else "存在差异，请目视对比 PDF", "====")
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(run_train())
