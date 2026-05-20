"""
GST Tax Invoice Module — Generates compliant tax invoices for every successful payment.

Behaviour:
  - GST-inclusive treatment: the amount paid by the customer is the FINAL total. The
    base amount and GST are reverse-calculated from the total at 18% GST.
  - Place of supply detection: if the buyer state == seller state (Rajasthan), GST is
    split as CGST 9% + SGST 9%. Otherwise, IGST 18% is applied.
  - Invoice numbering: MB/{FY}/{NNNN} (financial year, padded sequence) — atomic via
    a per-FY counter document in `invoice_counters`.
  - Stores every generated invoice metadata in the `invoices` collection.
  - PDF generation is portrait A4 with full GST-compliant header (seller GSTIN,
    address) and amount-in-words.

Public entry points used by server.py / cma_module.py:

    from invoice_module import create_invoice_for_payment, generate_invoice_pdf, register_invoice_routes

`create_invoice_for_payment(...)` is awaitable and must be called *after* the
payment status has been moved to "paid". It is idempotent — calling it twice for
the same (kind, ref_id) returns the existing invoice.
"""
from __future__ import annotations

import io
import uuid
from datetime import datetime, timezone
from typing import Any, Dict, Optional

from fastapi import APIRouter, Depends, HTTPException
from fastapi.responses import StreamingResponse

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import (
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)


# Constants
GST_RATE = 18.0  # percent — fixed for SAC 998314 (Management consulting)
SAC_CODE = "998314"  # Management consultancy & related services
ITEM_DESCRIPTIONS = {
    "dpr": "Detailed Project Report (DPR) — Bank-ready PDF + Excel",
    "cma": "CMA Data Preparation (Credit Monitoring Arrangement) — PDF + Excel",
    "wallet_topup": "Wallet Top-up — DPRForge",
    "custom": "DPRForge — Professional Services",
}

# Indian states for GST place-of-supply detection
_NORMALIZED_RAJASTHAN = {"rajasthan", "raj", "rj"}


def _is_intra_state(buyer_state: Optional[str], seller_state: str) -> bool:
    """Return True iff CGST+SGST applies (buyer in same state as seller)."""
    if not buyer_state:
        # Conservative default: treat as intra-state (CGST+SGST) so total still adds up.
        return True
    b = buyer_state.strip().lower()
    s = seller_state.strip().lower()
    return b == s or (b in _NORMALIZED_RAJASTHAN and s in _NORMALIZED_RAJASTHAN)


def _fy_label(dt: datetime) -> str:
    """Return Indian financial year label like '2025-26' for a UTC datetime."""
    y = dt.year
    # FY starts April 1
    if dt.month >= 4:
        return f"{y}-{str(y + 1)[-2:]}"
    return f"{y - 1}-{str(y)[-2:]}"


def _amount_in_words_inr(amount: float) -> str:
    """Convert a rupee amount to Indian-format words (lakhs / crores)."""
    # Convert to total paise as int first to avoid float rounding bugs
    total_paise = int(round(float(amount) * 100))
    n = total_paise // 100
    paise = total_paise % 100
    if n == 0:
        words = "Zero"
    else:
        words = _num_to_words_indian(n)
    out = f"INR {words} Rupees"
    if paise:
        out += f" and {_num_to_words_indian(paise)} Paise"
    return out + " Only"


_ONES = [
    "", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine",
    "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen",
    "Seventeen", "Eighteen", "Nineteen",
]
_TENS = ["", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"]


def _two_digit(n: int) -> str:
    if n < 20:
        return _ONES[n]
    return (_TENS[n // 10] + (" " + _ONES[n % 10] if n % 10 else "")).strip()


def _three_digit(n: int) -> str:
    h = n // 100
    rem = n % 100
    out = ""
    if h:
        out = _ONES[h] + " Hundred"
        if rem:
            out += " "
    if rem:
        out += _two_digit(rem)
    return out


def _num_to_words_indian(n: int) -> str:
    if n == 0:
        return "Zero"
    parts = []
    crore = n // 10000000
    n %= 10000000
    lakh = n // 100000
    n %= 100000
    thousand = n // 1000
    n %= 1000
    hundred = n
    if crore:
        parts.append(_two_digit(crore) + " Crore")
    if lakh:
        parts.append(_two_digit(lakh) + " Lakh")
    if thousand:
        parts.append(_two_digit(thousand) + " Thousand")
    if hundred:
        parts.append(_three_digit(hundred))
    return " ".join(parts)


# ---------------------------------------------------------------------------
# Invoice number generator (atomic per financial year)
# ---------------------------------------------------------------------------

async def _next_invoice_number(db, fy: str) -> str:
    counter_id = f"invoice_seq::{fy}"
    res = await db.invoice_counters.find_one_and_update(
        {"_id": counter_id},
        {"$inc": {"seq": 1}},
        upsert=True,
        return_document=True,
    )
    # find_one_and_update with upsert may return None on first insert in older drivers — re-fetch
    if not res:
        res = await db.invoice_counters.find_one({"_id": counter_id})
    seq = (res or {}).get("seq") or 1
    return f"MB/{fy}/{int(seq):04d}"


# ---------------------------------------------------------------------------
# Core: create invoice (idempotent)
# ---------------------------------------------------------------------------

# Ensure idempotency under concurrent calls using a unique index. Created
# once per process (cached via _indexes_ready).
_indexes_ready = False


async def _ensure_indexes(db) -> None:
    global _indexes_ready
    if _indexes_ready:
        return
    try:
        await db.invoices.create_index([("kind", 1), ("ref_id", 1)], unique=True, name="uniq_kind_ref")
        await db.invoices.create_index([("buyer_user_id", 1), ("created_at", -1)], name="user_recent")
    except Exception:
        # Index creation is best-effort; duplicate / pre-existing index is fine.
        pass
    _indexes_ready = True


async def create_invoice_for_payment(
    db,
    *,
    kind: str,                     # 'dpr' | 'cma' | 'wallet_topup' | 'custom'
    ref_id: str,                   # project_id / cma_id / order_id
    user: Dict[str, Any],          # full user doc (must have user_id, email, name)
    amount_paid: float,            # GST-INCLUSIVE total the customer paid
    payment_method: str,
    payment_txn_id: str,
    seller: Dict[str, Any],        # seller info (name, address, state, gstin, etc.)
    buyer_state: Optional[str] = None,
    buyer_name: Optional[str] = None,
    buyer_address: Optional[str] = None,
    description: Optional[str] = None,
) -> Dict[str, Any]:
    """Idempotently create (or fetch) a tax invoice record for a paid transaction.

    Returns the invoice document (without `_id`).
    """
    await _ensure_indexes(db)
    # Idempotency: if invoice already exists for this (kind, ref_id), return it.
    existing = await db.invoices.find_one({"kind": kind, "ref_id": ref_id}, {"_id": 0})
    if existing:
        return existing

    now = datetime.now(timezone.utc)
    fy = _fy_label(now)
    invoice_no = await _next_invoice_number(db, fy)

    # GST-inclusive: reverse-calculate base + tax
    total = round(float(amount_paid or 0), 2)
    base = round(total / (1 + GST_RATE / 100.0), 2)
    gst_total = round(total - base, 2)
    intra = _is_intra_state(buyer_state, seller.get("state", "Rajasthan"))
    if intra:
        cgst = round(gst_total / 2.0, 2)
        sgst = round(gst_total - cgst, 2)
        igst = 0.0
    else:
        cgst = 0.0
        sgst = 0.0
        igst = gst_total

    invoice = {
        "invoice_id": str(uuid.uuid4()),
        "invoice_no": invoice_no,
        "invoice_date": now.isoformat(),
        "fy": fy,
        "kind": kind,
        "ref_id": ref_id,

        # Seller
        "seller_name": seller.get("name", "Mother Bless Digital Solutions"),
        "seller_address_line1": seller.get("address_line1", ""),
        "seller_address_line2": seller.get("address_line2", ""),
        "seller_city": seller.get("city", ""),
        "seller_state": seller.get("state", "Rajasthan"),
        "seller_pincode": seller.get("pincode", ""),
        "seller_country": seller.get("country", "India"),
        "seller_gstin": seller.get("gstin", "08KQRPS8229A1Z6"),
        "seller_email": seller.get("email", ""),
        "seller_phone": seller.get("primary_phone", ""),

        # Buyer
        "buyer_user_id": (user or {}).get("user_id"),
        "buyer_name": buyer_name or (user or {}).get("name", "Customer"),
        "buyer_email": (user or {}).get("email", ""),
        "buyer_address": buyer_address or "",
        "buyer_state": buyer_state or "",

        # Item
        "description": description or ITEM_DESCRIPTIONS.get(kind, ITEM_DESCRIPTIONS["custom"]),
        "sac_code": SAC_CODE,
        "qty": 1,

        # Money (₹)
        "amount_paid": total,
        "taxable_value": base,
        "gst_rate": GST_RATE,
        "cgst_rate": 9.0 if intra else 0.0,
        "sgst_rate": 9.0 if intra else 0.0,
        "igst_rate": 0.0 if intra else 18.0,
        "cgst_amount": cgst,
        "sgst_amount": sgst,
        "igst_amount": igst,
        "total_tax": gst_total,
        "total_amount": total,
        "amount_in_words": _amount_in_words_inr(total),

        # Payment
        "payment_method": payment_method or "UPI",
        "payment_txn_id": payment_txn_id or "",
        "payment_date": now.isoformat(),

        "created_at": now.isoformat(),
    }

    try:
        await db.invoices.insert_one(dict(invoice))
    except Exception:
        # Race condition: another concurrent call inserted first. Return that one.
        existing = await db.invoices.find_one({"kind": kind, "ref_id": ref_id}, {"_id": 0})
        if existing:
            return existing
        raise
    invoice.pop("_id", None)
    return invoice


# ---------------------------------------------------------------------------
# PDF generation
# ---------------------------------------------------------------------------

def _fmt_inr(v: float) -> str:
    try:
        n = float(v)
    except (TypeError, ValueError):
        return "0.00"
    # Indian numbering (12,34,567.89)
    neg = n < 0
    n = abs(n)
    int_part = int(n)
    dec_part = round((n - int_part) * 100)
    s = str(int_part)
    if len(s) <= 3:
        body = s
    else:
        last3 = s[-3:]
        rest = s[:-3]
        # group rest in 2s
        groups = []
        while len(rest) > 2:
            groups.insert(0, rest[-2:])
            rest = rest[:-2]
        if rest:
            groups.insert(0, rest)
        body = ",".join(groups) + "," + last3
    out = f"{body}.{dec_part:02d}"
    return f"-{out}" if neg else out


def generate_invoice_pdf(invoice: Dict[str, Any]) -> bytes:
    """Render a GST-compliant tax invoice as a PDF (portrait A4)."""
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=15 * mm, rightMargin=15 * mm,
        topMargin=15 * mm, bottomMargin=15 * mm,
        title=f"Tax Invoice {invoice.get('invoice_no', '')}",
    )
    styles = getSampleStyleSheet()
    title = ParagraphStyle(
        "InvTitle", parent=styles["Title"], fontSize=18,
        textColor=colors.HexColor("#0F172A"), alignment=1, spaceAfter=2,
    )
    subtitle = ParagraphStyle(
        "InvSub", parent=styles["BodyText"], fontSize=9,
        textColor=colors.HexColor("#475569"), alignment=1, spaceAfter=4,
    )
    h_label = ParagraphStyle(
        "Lbl", parent=styles["BodyText"], fontSize=7.5,
        textColor=colors.HexColor("#64748B"), spaceAfter=1, leading=9,
    )
    h_val = ParagraphStyle(
        "Val", parent=styles["BodyText"], fontSize=9.5,
        textColor=colors.HexColor("#0F172A"), leading=12,
    )
    note = ParagraphStyle(
        "Note", parent=styles["BodyText"], fontSize=8,
        textColor=colors.HexColor("#475569"), leading=11,
    )

    story = []

    # --- Header band: seller block ---
    seller_lines = [
        f"<b>{invoice.get('seller_name', 'Mother Bless Digital Solutions')}</b>",
        " ".join([x for x in [invoice.get("seller_address_line1"), invoice.get("seller_address_line2")] if x]),
        ", ".join([x for x in [invoice.get("seller_city"), invoice.get("seller_state"), invoice.get("seller_pincode")] if x]),
        invoice.get("seller_country", "India"),
        f"GSTIN: <b>{invoice.get('seller_gstin', '')}</b>",
        f"Email: {invoice.get('seller_email', '') or '—'} &nbsp;&nbsp; Phone: {invoice.get('seller_phone', '') or '—'}",
    ]
    seller_block = "<br/>".join([s for s in seller_lines if s])

    story.append(Paragraph("TAX INVOICE", title))
    story.append(Paragraph("Original for Recipient", subtitle))
    story.append(Spacer(1, 3 * mm))

    header_tbl = Table(
        [[Paragraph(seller_block, h_val),
          Paragraph(
              f"<b>Invoice No.</b><br/>{invoice.get('invoice_no', '')}<br/><br/>"
              f"<b>Invoice Date</b><br/>{_fmt_date(invoice.get('invoice_date'))}<br/><br/>"
              f"<b>FY</b> {invoice.get('fy', '')}",
              h_val,
          )]],
        colWidths=[115 * mm, 60 * mm],
    )
    header_tbl.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F8FAFC")),
        ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
        ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))
    story.append(header_tbl)

    # --- Bill To ---
    bill_to_lines = [
        f"<b>{invoice.get('buyer_name', 'Customer')}</b>",
        invoice.get("buyer_email", "") or "",
        invoice.get("buyer_address", "") or "",
        f"State: {invoice.get('buyer_state', '') or '—'}",
    ]
    bill_to_block = "<br/>".join([s for s in bill_to_lines if s])
    intra = (invoice.get("igst_amount", 0) or 0) == 0
    pos_state = invoice.get("buyer_state") or invoice.get("seller_state", "Rajasthan")

    bill_tbl = Table(
        [
            [Paragraph("<b>BILL TO</b>", h_label),
             Paragraph("<b>PLACE OF SUPPLY</b>", h_label)],
            [Paragraph(bill_to_block, h_val),
             Paragraph(
                 f"{pos_state}<br/>"
                 f"<font size=8 color='#64748B'>{'Intra-state (CGST+SGST)' if intra else 'Inter-state (IGST)'}</font>",
                 h_val,
             )],
        ],
        colWidths=[115 * mm, 60 * mm],
    )
    bill_tbl.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0F172A")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
        ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))
    story.append(Spacer(1, 3 * mm))
    story.append(bill_tbl)

    # --- Items table ---
    story.append(Spacer(1, 4 * mm))

    if intra:
        items_header = [
            "S.No.", "Description", "SAC", "Qty",
            "Taxable Value", "CGST (9%)", "SGST (9%)", "Total",
        ]
        items_row = [
            "1",
            Paragraph(invoice.get("description", ""), h_val),
            invoice.get("sac_code", SAC_CODE),
            "1",
            _fmt_inr(invoice.get("taxable_value", 0)),
            _fmt_inr(invoice.get("cgst_amount", 0)),
            _fmt_inr(invoice.get("sgst_amount", 0)),
            _fmt_inr(invoice.get("total_amount", 0)),
        ]
        col_widths = [12 * mm, 60 * mm, 14 * mm, 10 * mm, 22 * mm, 19 * mm, 19 * mm, 19 * mm]
    else:
        items_header = [
            "S.No.", "Description", "SAC", "Qty",
            "Taxable Value", "IGST (18%)", "Total",
        ]
        items_row = [
            "1",
            Paragraph(invoice.get("description", ""), h_val),
            invoice.get("sac_code", SAC_CODE),
            "1",
            _fmt_inr(invoice.get("taxable_value", 0)),
            _fmt_inr(invoice.get("igst_amount", 0)),
            _fmt_inr(invoice.get("total_amount", 0)),
        ]
        col_widths = [12 * mm, 68 * mm, 14 * mm, 10 * mm, 25 * mm, 23 * mm, 23 * mm]

    items_tbl = Table([items_header, items_row], colWidths=col_widths, repeatRows=1)
    items_tbl.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0F172A")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8.5),
        ("ALIGN", (0, 0), (-1, 0), "CENTER"),
        ("ALIGN", (2, 1), (-1, -1), "RIGHT"),
        ("ALIGN", (0, 1), (0, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))
    story.append(items_tbl)

    # --- Totals block ---
    totals_rows = [
        ["Taxable Value", _fmt_inr(invoice.get("taxable_value", 0))],
    ]
    if intra:
        totals_rows.append(["CGST @ 9%", _fmt_inr(invoice.get("cgst_amount", 0))])
        totals_rows.append(["SGST @ 9%", _fmt_inr(invoice.get("sgst_amount", 0))])
    else:
        totals_rows.append(["IGST @ 18%", _fmt_inr(invoice.get("igst_amount", 0))])
    totals_rows.append(["Total Tax", _fmt_inr(invoice.get("total_tax", 0))])
    totals_rows.append(["Grand Total (INR)", _fmt_inr(invoice.get("total_amount", 0))])

    totals_tbl = Table(totals_rows, colWidths=[120 * mm, 55 * mm])
    totals_tbl.setStyle(TableStyle([
        ("ALIGN", (0, 0), (0, -1), "RIGHT"),
        ("ALIGN", (1, 0), (1, -1), "RIGHT"),
        ("FONTSIZE", (0, 0), (-1, -1), 9.5),
        ("TEXTCOLOR", (0, 0), (-1, -1), colors.HexColor("#0F172A")),
        ("LINEBELOW", (0, 0), (-1, -2), 0.25, colors.HexColor("#E2E8F0")),
        ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#0F172A")),
        ("TEXTCOLOR", (0, -1), (-1, -1), colors.white),
        ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("LEFTPADDING", (0, 0), (-1, -1), 10),
        ("RIGHTPADDING", (0, 0), (-1, -1), 10),
    ]))
    story.append(Spacer(1, 3 * mm))
    story.append(totals_tbl)

    # Amount in words
    story.append(Spacer(1, 3 * mm))
    story.append(Paragraph(
        f"<b>Amount in Words:</b> {invoice.get('amount_in_words', '')}", h_val,
    ))

    # Payment details
    pay_tbl = Table(
        [[
            Paragraph(
                f"<b>Payment Method</b><br/>{invoice.get('payment_method', '')}", h_val),
            Paragraph(
                f"<b>Transaction ID</b><br/>{invoice.get('payment_txn_id', '') or '—'}", h_val),
            Paragraph(
                f"<b>Payment Date</b><br/>{_fmt_date(invoice.get('payment_date'))}", h_val),
        ]],
        colWidths=[60 * mm, 60 * mm, 55 * mm],
    )
    pay_tbl.setStyle(TableStyle([
        ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
        ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F8FAFC")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))
    story.append(Spacer(1, 4 * mm))
    story.append(pay_tbl)

    # Footer / notes
    story.append(Spacer(1, 6 * mm))
    story.append(Paragraph(
        "<b>Notes:</b> This is a digitally generated tax invoice. The amount paid is "
        "inclusive of GST @ 18%. No signature required. Subject to Bagidora (Rajasthan) jurisdiction.",
        note,
    ))
    story.append(Spacer(1, 2 * mm))
    story.append(Paragraph(
        "Thank you for choosing DPRForge by Mother Bless Digital Solutions.", note,
    ))

    doc.build(story)
    return buf.getvalue()


def _fmt_date(iso_str: Optional[str]) -> str:
    if not iso_str:
        return ""
    try:
        if isinstance(iso_str, datetime):
            dt = iso_str
        else:
            dt = datetime.fromisoformat(iso_str.replace("Z", "+00:00"))
        return dt.strftime("%d %b %Y")
    except Exception:
        return str(iso_str)[:10]


# ---------------------------------------------------------------------------
# Public routes (mounted under /api/invoices/*)
# ---------------------------------------------------------------------------

def register_invoice_routes(api_router: APIRouter, db, get_current_user):
    """Attach invoice download/list endpoints to the main API router."""

    @api_router.get("/invoices/my")
    async def my_invoices(user=Depends(get_current_user)):
        rows = await db.invoices.find(
            {"buyer_user_id": user.user_id}, {"_id": 0},
        ).sort("created_at", -1).to_list(500)
        return rows

    @api_router.get("/invoices/{invoice_id}")
    async def get_invoice(invoice_id: str, user=Depends(get_current_user)):
        inv = await db.invoices.find_one(
            {"invoice_id": invoice_id}, {"_id": 0},
        )
        if not inv:
            raise HTTPException(status_code=404, detail="Invoice not found")
        if inv.get("buyer_user_id") != user.user_id and not getattr(user, "is_admin", False):
            raise HTTPException(status_code=403, detail="Forbidden")
        return inv

    @api_router.get("/invoices/{invoice_id}/download")
    async def download_invoice(invoice_id: str, user=Depends(get_current_user)):
        inv = await db.invoices.find_one(
            {"invoice_id": invoice_id}, {"_id": 0},
        )
        if not inv:
            raise HTTPException(status_code=404, detail="Invoice not found")
        if inv.get("buyer_user_id") != user.user_id and not getattr(user, "is_admin", False):
            raise HTTPException(status_code=403, detail="Forbidden")
        pdf_bytes = generate_invoice_pdf(inv)
        safe_no = inv.get("invoice_no", "invoice").replace("/", "_")
        filename = f"Invoice_{safe_no}.pdf"
        return StreamingResponse(
            io.BytesIO(pdf_bytes),
            media_type="application/pdf",
            headers={"Content-Disposition": f'attachment; filename="{filename}"'},
        )

    @api_router.get("/projects/{project_id}/invoice")
    async def get_invoice_for_project(project_id: str, user=Depends(get_current_user)):
        """Convenience: fetch invoice by project_id for the current user."""
        inv = await db.invoices.find_one(
            {"kind": "dpr", "ref_id": project_id, "buyer_user_id": user.user_id}, {"_id": 0},
        )
        if not inv:
            raise HTTPException(status_code=404, detail="No invoice yet for this project")
        return inv

    @api_router.get("/cma/statements/{cma_id}/invoice")
    async def get_invoice_for_cma(cma_id: str, user=Depends(get_current_user)):
        inv = await db.invoices.find_one(
            {"kind": "cma", "ref_id": cma_id, "buyer_user_id": user.user_id}, {"_id": 0},
        )
        if not inv:
            raise HTTPException(status_code=404, detail="No invoice yet for this CMA")
        return inv

    @api_router.get("/admin/invoices")
    async def admin_list_invoices(user=Depends(get_current_user), limit: int = 500):
        if not getattr(user, "is_admin", False):
            raise HTTPException(status_code=403, detail="Admin only")
        rows = await db.invoices.find({}, {"_id": 0}).sort("created_at", -1).to_list(limit)
        return rows
