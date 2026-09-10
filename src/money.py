"""Decimal money helpers — never use float for amount math."""

from __future__ import annotations

from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from typing import Any

TWOPLACES = Decimal("0.01")


def to_decimal(value: Any) -> Decimal | None:
    if value is None:
        return None
    if isinstance(value, Decimal):
        return value
    if isinstance(value, bool):
        return None
    if isinstance(value, int):
        return Decimal(value)
    if isinstance(value, float):
        # Convert via str to avoid binary float artifacts where possible.
        return Decimal(str(value))
    text = str(value).strip()
    if not text:
        return None
    text = text.replace(",", "").replace("$", "")
    try:
        return Decimal(text)
    except (InvalidOperation, ValueError):
        return None


def quantize(value: Decimal, places: str | Decimal = TWOPLACES) -> Decimal:
    q = Decimal(places) if not isinstance(places, Decimal) else places
    return value.quantize(q, rounding=ROUND_HALF_UP)


def money_diff(a: Decimal, b: Decimal, places: str | Decimal = TWOPLACES) -> Decimal:
    return abs(quantize(a, places) - quantize(b, places))


def format_money(value: Decimal) -> str:
    q = quantize(value)
    sign = "-" if q < 0 else ""
    abs_q = abs(q)
    return f"{sign}${abs_q:,.2f}"
