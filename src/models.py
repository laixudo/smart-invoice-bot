from __future__ import annotations

from dataclasses import dataclass, field
from decimal import Decimal

from .money import quantize


@dataclass
class LineItem:
    cbm: Decimal | None
    destination: str
    isa: str
    qty: Decimal | None
    rate: Decimal | None
    amount: Decimal
    source_row: int


@dataclass
class ContainerInvoice:
    container_no: str
    total_volume: Decimal | None
    invoice_no: str
    bill_name: str
    invoice_date: str
    bill_to: str
    source_final_settlement: Decimal | None = None
    source_final_settlement_row: int | None = None
    source_feedback: str = ""
    source_feedback_row: int | None = None
    lines: list[LineItem] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    @property
    def total_amount(self) -> Decimal:
        """Customer-approved amount: add raw line amounts per container, then round.

        PDF line items still display two decimals, but the invoice total is derived
        from the full-precision sum of that container's source lines and rounded
        only once at the end.
        """
        raw_total = sum((ln.amount for ln in self.lines), Decimal("0"))
        return quantize(raw_total)

    @property
    def display_lines(self) -> list[LineItem]:
        """The PDF contains only real source lines; no balancing line is used."""
        return list(self.lines)


@dataclass
class QuarantineRow:
    row_number: int
    reason: str
    raw: dict


@dataclass
class ContainerValidation:
    container_no: str
    status: str  # OK | WARN | FAIL
    system_total: Decimal
    summary_total: Decimal | None
    diff: Decimal | None
    message: str = ""


@dataclass
class BatchResult:
    invoices: list[ContainerInvoice]
    validations: list[ContainerValidation]
    quarantine: list[QuarantineRow]
    batch_ok: bool
    batch_message: str
    bill_name: str
    invoice_date: str
    summary_grand_total: Decimal | None
    system_grand_total: Decimal
    schema_warnings: list[QuarantineRow] = field(default_factory=list)
    main_sheet_name: str = ""
