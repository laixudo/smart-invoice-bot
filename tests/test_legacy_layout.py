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


def _legacy_workbook(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "BasicInvoice"
    ws.append(
        [
            "APPOINTMENT ID",
            "DESTINATION",
            "ISA",
            "QTY",
            "RATE",
            "AMOUNT",
            "合计金额",
            "invoice编号",
            None,
            "最终结算金额",
            "AR发票",
        ]
    )
    ws.append(
        [
            "ABCD1234567-45HC  80.5",
            "LAX9 - 10.5",
            "ISA-1",
            Decimal("0.5"),
            100,
            100,
            100,
            2603001,
            None,
            "OK",
            '=DISPIMG("IMAGE_ID",1)',
        ]
    )
    ws.append(
        [
            None,
            "SBD1-GY20.0",
            "ISA-2",
            Decimal("0.5"),
            100,
            None,
            None,
            None,
            None,
            None,
            None,
        ]
    )
    ws.append([None, None, None, None, "Total Amount:", 100])
    for column in ("A", "F", "G", "H", "J", "K"):
        ws.merge_cells(f"{column}2:{column}3")
    wb.save(path)


class LegacyLayoutTests(unittest.TestCase):
    def test_legacy_mixed_columns_are_normalized_without_changing_total(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "#DS-03042026-SMART.xlsx"
            _legacy_workbook(path)

            parsed = parse_excel(path, CFG)
            batch = validate_batch(parsed, CFG)

            self.assertTrue(batch.batch_ok, batch.batch_message)
            self.assertEqual(len(parsed.invoices), 1)
            invoice = parsed.invoices[0]
            self.assertEqual(invoice.container_no, "ABCD1234567")
            self.assertEqual(invoice.total_volume, Decimal("80.5"))
            self.assertEqual(invoice.invoice_no, "2603001")
            self.assertEqual(invoice.total_amount, Decimal("100.00"))
            self.assertEqual([line.destination for line in invoice.lines], ["LAX9", "SBD1"])
            self.assertEqual([line.cbm for line in invoice.lines], [Decimal("10.5"), Decimal("20.0")])
            self.assertEqual([line.amount for line in invoice.lines], [Decimal("100"), Decimal("0")])
            self.assertFalse(
                any(w.reason == "最终结算金额非数字" for w in parsed.schema_warnings)
            )

    def test_missing_legacy_total_volume_falls_back_to_cbm_sum(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "#DS-03042026-SMART.xlsx"
            wb = Workbook()
            ws = wb.active
            ws.title = "BasicInvoice"
            ws.append(
                [
                    "APPOINTMENT ID", "DESTINATION", "ISA", "QTY", "RATE", "AMOUNT",
                    "合计金额", "invoice编号", None, "最终结算金额", "AR发票",
                ]
            )
            ws.append(["EFGH1234567", "GEU2 - 3.25", "ISA-3", 1, 50, 50, 50, 2603002])
            wb.save(path)

            parsed = parse_excel(path, CFG)

            self.assertEqual(parsed.invoices[0].total_volume, Decimal("3.25"))
            self.assertTrue(
                any(w.reason == "旧版模板总体积由明细CBM合计" for w in parsed.schema_warnings)
            )

    def test_invoice_number_is_backfilled_after_cross_boundary_merge(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "#DS-03042026-SMART.xlsx"
            wb = Workbook()
            ws = wb.active
            ws.title = "BasicInvoice"
            ws.append(
                [
                    "APPOINTMENT ID", "DESTINATION", "ISA", "QTY", "RATE", "AMOUNT",
                    "合计金额", "invoice编号", None, "最终结算金额", "AR发票",
                ]
            )
            ws.append(["ABCD1234567", "LAX9 - 1", "ISA-1", 1, 10, 10, 10, 2603001])
            ws.append([None, "私人地址-A", "CA 91745", 1, 230, 230])
            ws.append(["EFGH1234567", "私人地址-B", "CA 91789", None, None, None, 230, 2603013])
            ws.merge_cells("A2:A3")
            ws.merge_cells("D3:D4")
            ws.merge_cells("E3:E4")
            ws.merge_cells("F3:F4")
            wb.save(path)

            parsed = parse_excel(path, CFG)
            by_id = {invoice.container_no: invoice for invoice in parsed.invoices}

            self.assertEqual(by_id["ABCD1234567"].total_amount, Decimal("10.00"))
            self.assertEqual(by_id["EFGH1234567"].total_amount, Decimal("230.00"))
            self.assertEqual(by_id["EFGH1234567"].invoice_no, "2603013")


if __name__ == "__main__":
    unittest.main()
