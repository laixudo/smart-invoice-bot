from __future__ import annotations

from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    BaseDocTemplate,
    Frame,
    PageTemplate,
    Table,
    TableStyle,
    Paragraph,
    Spacer,
    PageBreak,
)

from .models import ContainerInvoice
from .money import format_money, quantize


_FONT_REGISTERED = False
_FONT_NAME = "Helvetica"
_FONT_BOLD = "Helvetica-Bold"

PAGE_SIZE = landscape(A4)  # 对齐黄金样例 841.89 x 595.28
GRAY = colors.HexColor("#808080")
DARK = colors.HexColor("#222222")


def _ensure_font(chinese_font_path: str | None) -> None:
    global _FONT_REGISTERED, _FONT_NAME, _FONT_BOLD
    if _FONT_REGISTERED:
        return
    path = Path(chinese_font_path or "")
    if path.exists():
        pdfmetrics.registerFont(TTFont("InvoiceCN", str(path)))
        _FONT_NAME = "InvoiceCN"
        _FONT_BOLD = "InvoiceCN"
    _FONT_REGISTERED = True


def _fmt_num(value, max_decimals: int = 8) -> str:
    if value is None:
        return ""
    if max_decimals == 2:
        q = quantize(value, "0.01")
        if q == q.to_integral():
            return f"{int(q)}"
        return f"{q:.2f}"
    exp = Decimal("1").scaleb(-max_decimals)
    q = Decimal(value).quantize(exp, rounding=ROUND_HALF_UP)
    text = f"{q:f}"
    if "." in text:
        text = text.rstrip("0").rstrip(".")
    return text


def render_invoice_pdf(invoice: ContainerInvoice, out_path: Path, cfg: dict) -> Path:
    company = cfg.get("company", {})
    _ensure_font(cfg.get("chinese_font_path"))
    # Keep a conservative single-page cap so totals and the contact footer never spill onto a blank page.
    max_lines = min(int(cfg.get("max_lines_single_page", 15)), 13)
    min_font = int(cfg.get("min_font_size", 7))

    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    lines = invoice.display_lines
    n = len(lines)
    if n <= max_lines:
        body_font = 9 if n <= 12 else max(min_font, 11 - (n - 10))
        multipage = False
    else:
        body_font = max(min_font, 8)
        multipage = True

    page_w, page_h = PAGE_SIZE
    left = 12 * mm
    right = 12 * mm
    top = 10 * mm
    bottom = 10 * mm
    content_w = page_w - left - right

    doc = BaseDocTemplate(
        str(out_path),
        pagesize=PAGE_SIZE,
        leftMargin=left,
        rightMargin=right,
        topMargin=top,
        bottomMargin=bottom,
    )
    frame = Frame(left, bottom, content_w, page_h - top - bottom, id="normal")

    def _on_page(canvas, doc_):
        if multipage:
            canvas.setFont(_FONT_NAME, 8)
            canvas.setFillColor(colors.HexColor("#666666"))
            canvas.drawRightString(page_w - right, 6 * mm, f"Page {canvas.getPageNumber()}")

    doc.addPageTemplates([PageTemplate(id="main", frames=[frame], onPage=_on_page)])

    s_company = ParagraphStyle(
        "company", fontName=_FONT_BOLD, fontSize=13, leading=16, textColor=DARK
    )
    s_small = ParagraphStyle(
        "small", fontName=_FONT_NAME, fontSize=9, leading=12, textColor=DARK
    )
    s_invoice = ParagraphStyle(
        "invoice_title",
        fontName=_FONT_BOLD,
        fontSize=32,
        leading=36,
        textColor=colors.HexColor("#9a9a9a"),
        alignment=2,
    )
    s_bill_name = ParagraphStyle(
        "bill_name",
        fontName=_FONT_BOLD,
        fontSize=12,
        leading=14,
        textColor=DARK,
        alignment=2,
    )
    s_th = ParagraphStyle(
        "th", fontName=_FONT_BOLD, fontSize=8, leading=10, textColor=colors.white, alignment=1
    )
    s_td = ParagraphStyle(
        "td", fontName=_FONT_NAME, fontSize=body_font, leading=body_font + 2, alignment=1
    )
    s_td_left = ParagraphStyle(
        "td_left", fontName=_FONT_NAME, fontSize=body_font, leading=body_font + 2, alignment=0
    )

    left_block = [
        Paragraph(company.get("name", ""), s_company),
        *[Paragraph(x, s_small) for x in company.get("address_lines", [])],
        Paragraph(f"Phone: {company.get('phone', '')}", s_small),
        Paragraph(f"Account Number ：{company.get('account_number', '')}", s_small),
        Paragraph(f"Routing Number：{company.get('routing_number', '')}", s_small),
        Paragraph(f"Zelle ：{company.get('zelle', '')}", s_small),
        Paragraph(company.get("emails", ""), s_small),
    ]

    date_box = Table(
        [
            [
                Paragraph(
                    "DATE",
                    ParagraphStyle(
                        "dh",
                        fontName=_FONT_BOLD,
                        fontSize=9,
                        textColor=colors.white,
                        alignment=1,
                    ),
                )
            ],
            [
                Paragraph(
                    invoice.invoice_date,
                    ParagraphStyle("dv", fontName=_FONT_NAME, fontSize=11, alignment=1),
                )
            ],
        ],
        colWidths=[42 * mm],
    )
    date_box.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), GRAY),
                ("BOX", (0, 0), (-1, -1), 0.7, colors.HexColor("#666666")),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ]
        )
    )

    right_block = [
        Paragraph("INVOICE", s_invoice),
        Spacer(1, 2 * mm),
        Paragraph(invoice.bill_name, s_bill_name),
        Spacer(1, 4 * mm),
        date_box,
    ]

    header = Table([[left_block, right_block]], colWidths=[content_w * 0.62, content_w * 0.38])
    header.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))

    bill_bar = Table(
        [
            [
                Paragraph(
                    "BILL TO",
                    ParagraphStyle(
                        "bt", fontName=_FONT_BOLD, fontSize=11, textColor=colors.white
                    ),
                )
            ]
        ],
        colWidths=[content_w],
    )
    bill_bar.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), GRAY),
                ("LEFTPADDING", (0, 0), (-1, -1), 8),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ]
        )
    )

    inv_no = Table(
        [
            [
                "",
                Paragraph(
                    f"Invoice no： {invoice.invoice_no}",
                    ParagraphStyle("ino", fontName=_FONT_NAME, fontSize=11, alignment=2),
                ),
            ]
        ],
        colWidths=[content_w * 0.65, content_w * 0.35],
    )

    headers = [
        "APPOINTMENT ID",
        "Total Volume",
        "CBM",
        "DESTINATION",
        "ISA",
        "QTY",
        "RATE",
        "AMOUNT",
    ]
    vol_text = _fmt_num(invoice.total_volume)

    # Each PDF page owns one complete table section.  That lets the first two
    # columns be vertically spanned on every page, including continuation pages.
    col_widths = [
        content_w * 0.14,
        content_w * 0.10,
        content_w * 0.08,
        content_w * 0.14,
        content_w * 0.18,
        content_w * 0.12,
        content_w * 0.10,
        content_w * 0.14,
    ]

    def _make_invoice_table(page_lines):
        data = [[Paragraph(h, s_th) for h in headers]]
        for i, ln in enumerate(page_lines):
            data.append(
                [
                    Paragraph(invoice.container_no if i == 0 else "", s_td_left),
                    Paragraph(vol_text if i == 0 else "", s_td),
                    Paragraph(_fmt_num(ln.cbm), s_td),
                    Paragraph(ln.destination, s_td),
                    Paragraph(ln.isa, s_td),
                    Paragraph(_fmt_num(ln.qty, 8), s_td),
                    Paragraph(_fmt_num(ln.rate, 2), s_td),
                    Paragraph(_fmt_num(ln.amount, 2), s_td),
                ]
            )

        table = Table(data, colWidths=col_widths, repeatRows=0)
        styles = [
            ("BACKGROUND", (0, 0), (-1, 0), GRAY),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#777777")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ("LEFTPADDING", (0, 0), (-1, -1), 3),
            ("RIGHTPADDING", (0, 0), (-1, -1), 3),
            (
                "ROWBACKGROUNDS",
                (0, 1),
                (-1, -1),
                [colors.white, colors.HexColor("#f5f5f5")],
            ),
        ]
        if len(page_lines) > 1:
            styles.extend(
                [
                    ("SPAN", (0, 1), (0, len(page_lines))),
                    ("SPAN", (1, 1), (1, len(page_lines))),
                    ("VALIGN", (0, 1), (1, len(page_lines)), "MIDDLE"),
                ]
            )
        table.setStyle(TableStyle(styles))
        return table

    page_chunks = [lines[i : i + max_lines] for i in range(0, len(lines), max_lines)]
    invoice_tables = [_make_invoice_table(chunk) for chunk in page_chunks]

    total_txt = format_money(invoice.total_amount)
    total_bar = Table(
        [
            [
                "",
                Paragraph(
                    f"Total Amount: {total_txt}",
                    ParagraphStyle(
                        "tot",
                        fontName=_FONT_BOLD,
                        fontSize=12,
                        textColor=colors.white,
                        alignment=2,
                    ),
                ),
            ]
        ],
        colWidths=[content_w * 0.62, content_w * 0.38],
    )
    total_bar.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (1, 0), (1, 0), GRAY),
                ("LEFTPADDING", (1, 0), (1, 0), 10),
                ("RIGHTPADDING", (1, 0), (1, 0), 10),
                ("TOPPADDING", (0, 0), (-1, -1), 7),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
            ]
        )
    )

    story = [
        header,
        Spacer(1, 5 * mm),
        bill_bar,
        Paragraph(
            invoice.bill_to,
            ParagraphStyle("bto", fontName=_FONT_BOLD, fontSize=13, leading=18),
        ),
        Spacer(1, 3 * mm),
        inv_no,
        Spacer(1, 3 * mm),
        invoice_tables[0],
    ]
    for continuation_table in invoice_tables[1:]:
        story.extend([PageBreak(), continuation_table])
    story.extend(
        [
            Spacer(1, 4 * mm),
            total_bar,
            Spacer(1, 2 * mm),
            Paragraph(
                company.get("contact_note", ""),
                ParagraphStyle("cn", fontName=_FONT_NAME, fontSize=10, alignment=1),
            ),
            Paragraph(
                company.get("contact_person", ""),
                ParagraphStyle("cp", fontName=_FONT_BOLD, fontSize=11, alignment=1),
            ),
        ]
    )

    doc.build(story)
    return out_path
