# -*- coding: utf-8 -*-
from __future__ import annotations

import sys
from decimal import Decimal
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from src.config_loader import load_config
from src.excel_parser import parse_excel
from src.money import quantize
from src.validator import validate_batch

SAMPLES = ROOT / "samples"


def _sample_xlsx() -> Path:
    for p in SAMPLES.glob("*.xlsx"):
        if "SMART" in p.name.upper() or p.name.startswith("#"):
            return p
    raise FileNotFoundError(f"no sample xlsx in {SAMPLES}")


def test_tiiu_and_batch():
    sample = _sample_xlsx()
    cfg = load_config()
    parsed = parse_excel(sample, cfg)
    by_id = {i.container_no: i for i in parsed.invoices}

    assert "TIIU6724883" in by_id
    ti = by_id["TIIU6724883"]
    assert len(ti.lines) == 5
    assert quantize(ti.total_amount) == Decimal("4862.71")
    assert ti.invoice_no == "20260601"
    assert ti.bill_name == "#DS-06052026-SMART"
    assert ti.invoice_date == "06/05/2026"

    batch = validate_batch(parsed, cfg)
    assert batch.batch_ok, batch.batch_message
    assert batch.system_grand_total == Decimal("67788.41")
    assert len(parsed.invoices) == 17


def test_feedback_file_applies_configured_chassis_and_pallet_rules():
    parsed = parse_excel(ROOT / "input" / "#DS-06292026-SMART（SAV)已确认153367.21.xlsx-问题反馈.xlsx", load_config())
    by_id = {i.container_no: i for i in parsed.invoices}
    chassis = next(line for line in by_id["CCLU7371371"].lines if line.destination == "车架费")
    assert chassis.amount == Decimal("40")
    pallet = next(line for line in by_id["CSNU8041789"].lines if line.qty == Decimal("20"))
    assert pallet.amount == Decimal("2000")


def test_total_follows_container_sum_then_rounding():
    parsed = parse_excel(ROOT / "input" / "#DS-06292026-SMART（SAV)已确认153367.21.xlsx-问题反馈.xlsx", load_config())
    inv = next(i for i in parsed.invoices if i.container_no == "BEAU6086740")
    assert quantize(inv.total_amount) == Decimal("3054.20")
    assert sum((quantize(line.amount) for line in inv.display_lines), Decimal("0")) == Decimal("3054.20")
    assert len(inv.display_lines) == len(inv.lines)
    assert sum((quantize(line.amount) for line in inv.display_lines), Decimal("0")) == Decimal("3054.21")


def test_fee_rules_audit_but_never_overwrite_source_amount():
    parsed = parse_excel(ROOT / "input" / "#DS-06292026-SMART（SAV)已确认153367.21.xlsx-问题反馈.xlsx", load_config())
    by_id = {i.container_no: i for i in parsed.invoices}
    chassis = next(line for line in by_id["CSGU6052147"].lines if line.destination == "车架费")
    assert chassis.amount == Decimal("0")
    assert any(w.reason == "费用规则与源表金额不一致" for w in parsed.schema_warnings)


if __name__ == "__main__":
    test_tiiu_and_batch()
    print("OK")
