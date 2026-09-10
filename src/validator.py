from __future__ import annotations

from decimal import Decimal

from .excel_parser import ParseOutput
from .money import money_diff, quantize, to_decimal
from .models import BatchResult, ContainerValidation


def validate_batch(parsed: ParseOutput, cfg: dict) -> BatchResult:
    warn_tol = to_decimal(cfg.get("container_warn_tolerance", "0.01")) or Decimal("0.01")
    fail_tol = to_decimal(cfg.get("container_fail_tolerance", "0.01")) or Decimal("0.01")
    batch_tol = to_decimal(cfg.get("batch_tolerance", "0.01")) or Decimal("0.01")
    places = cfg.get("amount_quantize", "0.01")

    validations: list[ContainerValidation] = []
    system_grand = Decimal("0")

    for inv in parsed.invoices:
        sys_total = quantize(inv.total_amount, places)
        system_grand += sys_total
        summary = parsed.summary_totals.get(inv.container_no)
        summary_q = quantize(summary, places) if summary is not None else None
        if summary_q is None:
            validations.append(
                ContainerValidation(
                    container_no=inv.container_no,
                    status="WARN",
                    system_total=sys_total,
                    summary_total=None,
                    diff=None,
                    message="汇总表无此柜号",
                )
            )
            inv.warnings.append("汇总表无此柜号")
            continue

        diff = money_diff(sys_total, summary_q, places)
        if diff > fail_tol:
            status = "FAIL"
            msg = f"单柜差额 {diff} 超过失败阈值 {fail_tol}"
        elif diff > 0 and diff <= warn_tol:
            status = "WARN"
            msg = f"单柜差额 {diff} 在容差内"
            inv.warnings.append(msg)
        else:
            status = "OK"
            msg = ""
        validations.append(
            ContainerValidation(
                container_no=inv.container_no,
                status=status,
                system_total=sys_total,
                summary_total=summary_q,
                diff=diff,
                message=msg,
            )
        )

    system_grand = quantize(system_grand, places)
    summary_grand = (
        quantize(parsed.summary_grand_total, places)
        if parsed.summary_grand_total is not None
        else None
    )

    batch_ok = True
    messages: list[str] = []

    if any(v.status == "FAIL" for v in validations):
        batch_ok = False
        messages.append("存在单柜 FAIL")

    if summary_grand is not None:
        batch_diff = money_diff(system_grand, summary_grand, places)
        if batch_diff > batch_tol and cfg.get("batch_block_on_total_mismatch", True):
            batch_ok = False
            messages.append(
                f"批次总计不平: system={system_grand} summary={summary_grand} diff={batch_diff}"
            )

    if any(q.reason == "最终结算金额非数字" for q in parsed.schema_warnings):
        batch_ok = False
        messages.append("最终结算金额存在非数字值")

    if cfg.get("batch_block_on_container_count_mismatch", True) and parsed.summary_totals:
        sys_set = {i.container_no for i in parsed.invoices}
        sum_set = set(parsed.summary_totals)
        if sys_set != sum_set:
            batch_ok = False
            missing = sorted(sum_set - sys_set)
            extra = sorted(sys_set - sum_set)
            messages.append(f"柜号集合不一致 missing={missing} extra={extra}")

    return BatchResult(
        invoices=parsed.invoices,
        validations=validations,
        quarantine=parsed.quarantine,
        batch_ok=batch_ok,
        batch_message="; ".join(messages) if messages else "OK",
        bill_name=parsed.bill_name,
        invoice_date=parsed.invoice_date,
        summary_grand_total=summary_grand,
        system_grand_total=system_grand,
        schema_warnings=parsed.schema_warnings,
        main_sheet_name=parsed.main_sheet_name,
    )
