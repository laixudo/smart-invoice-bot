# -*- coding: utf-8 -*-
from __future__ import annotations

import sys
import tempfile
import unittest
from decimal import Decimal
from pathlib import Path

from openpyxl import Workbook

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from src.excel_parser import parse_excel
from src.validator import validate_batch


CFG = {
    "header_scan_rows": 20,
    "header_required_keywords": ["APPOINTMENT ID", "DESTINATION", "AMOUNT"],
    "core_columns": [
        "appointment_id",
        "total_volume",
        "cbm",
        "destination",
        "isa",
        "qty",
        "rate",
        "amount",
    ],
    "tail_block_keywords": ["Total Amount", "总计", "合计"],
    "schema_detection": {"enabled": True},
}


def _make_workbook(path: Path, *, hide_final_settlement: bool) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "#DS-07252026-SMART(CA)"
    ws.append(
        [
            "APPOINTMENT ID",
            "Total Volume",
            "CBM",
            "DESTINATION",
            "ISA",
            "QTY",
            "RATE",
            "AMOUNT",
            "备注",
            "最终结算金额",
            None,
            None,
            "AR发票",
        ]
    )
    ws.append(
        [
            "TGBU8838850",
            10,
            10,
            "LAX9",
            "ISA-1",
            1,
            100,
            100,
            None,
            "OK",
            None,
            None,
            20260709,
        ]
    )
    ws.column_dimensions["J"].hidden = hide_final_settlement
    # 核心列即使隐藏也必须继续参与解析。
    ws.column_dimensions["H"].hidden = True
    wb.save(path)


class HiddenFinalSettlementTests(unittest.TestCase):
    def test_hidden_auxiliary_column_does_not_block_batch(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "hidden.xlsx"
            _make_workbook(path, hide_final_settlement=True)

            parsed = parse_excel(path, CFG)
            batch = validate_batch(parsed, CFG)

            self.assertTrue(batch.batch_ok, batch.batch_message)
            self.assertEqual(len(parsed.invoices), 1)
            self.assertEqual(parsed.invoices[0].total_amount, Decimal("100.00"))
            self.assertIsNone(parsed.invoices[0].source_final_settlement)
            self.assertFalse(
                any(w.reason == "最终结算金额非数字" for w in parsed.schema_warnings)
            )

    def test_visible_non_numeric_final_settlement_still_blocks(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "visible.xlsx"
            _make_workbook(path, hide_final_settlement=False)

            parsed = parse_excel(path, CFG)
            batch = validate_batch(parsed, CFG)

            self.assertFalse(batch.batch_ok)
            self.assertIn("最终结算金额存在非数字值", batch.batch_message)


if __name__ == "__main__":
    unittest.main()
