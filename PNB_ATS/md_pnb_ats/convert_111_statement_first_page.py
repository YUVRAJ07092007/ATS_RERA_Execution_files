"""Convert 111_PNB_Statement_First_Page_Subvention_Yes_22062026.pdf to markdown."""
import re
from pathlib import Path

import fitz

PNB_ROOT = Path(__file__).resolve().parents[1]
PDF = PNB_ROOT / "attachments" / "111_PNB_Statement_First_Page_Subvention_Yes_22062026.pdf"
OUT = Path(__file__).resolve().parent / "111_PNB_Statement_First_Page_Subvention_Yes_22062026.md"
SOURCE_NAME = "111_PNB_Statement_First_Page_Subvention_Yes_22062026.pdf"


def _after_label(text: str, label: str) -> str | None:
    pattern = re.compile(
        rf"{re.escape(label)}\s*:?\s*\n?\s*([^\n]+)",
        re.IGNORECASE,
    )
    match = pattern.search(text)
    return match.group(1).strip() if match else None


def _field_after(text: str, label: str) -> str | None:
    pattern = re.compile(
        rf"{re.escape(label)}\s*\n\s*([^\n]+)",
        re.IGNORECASE,
    )
    match = pattern.search(text)
    return match.group(1).strip() if match else None


def extract_key_details(page_text: str) -> list[tuple[str, str]]:
    statement_date = _after_label(page_text, "Date")
    period = _after_label(page_text, "Statement for the period")
    customer_id = _after_label(page_text, "Customer ID")
    loan_account = _after_label(page_text, "Loan Account Number")
    sanction_date = _field_after(page_text, "Sanction date")
    sanction_amount = _field_after(page_text, "Sanction loan amount")
    disbursed = _field_after(page_text, "Disbursed Amount")
    instalment = _field_after(page_text, "Current instalment amount")
    if not instalment:
        match = re.search(r"Current instalment amount\s+([\d,]+\.?\d*)", page_text)
        instalment = match.group(1) if match else None
    rate = _field_after(page_text, "Current rate of interest")
    apr = _field_after(page_text, "Annual Percentage \nRate(APR)")
    if not apr:
        apr = _field_after(page_text, "Annual Percentage Rate(APR)")
    subvention = _field_after(page_text, "Subvention")
    repayment = _field_after(page_text, "Repayment mode")
    loan_status = _field_after(page_text, "Loan status (Active/ Closed)")
    disbursement_type = _field_after(page_text, "Disbursement Type")
    balance_tenure = _field_after(page_text, "Balance loan tenure")
    charges_due = _field_after(page_text, "Charges/ Fee Due as on \ndate (Excluding Tax)")
    if not charges_due:
        charges_due = _field_after(page_text, "Charges/ Fee Due as on date (Excluding Tax)")
    excess = _field_after(page_text, "Excess Amount")
    principal_paid = _field_after(page_text, "Total Principal Paid")
    interest_paid = _field_after(page_text, "Total Interest Paid")

    rows: list[tuple[str, str]] = []
    if customer_id:
        rows.append(("**Customer ID**", customer_id))
    if loan_account:
        rows.append(("**Loan Account No.**", loan_account))
    if statement_date:
        rows.append(("**Statement date**", statement_date))
    if period:
        rows.append(("**Statement period**", period))
    if sanction_date:
        rows.append(("**Sanction date**", sanction_date))
    if sanction_amount:
        rows.append(("**Sanction amount**", f"Rs. {sanction_amount}"))
    if disbursed:
        rows.append(("**Disbursed (as on statement)**", f"Rs. {disbursed}"))
    if instalment:
        rows.append(("**Current instalment**", f"Rs. {instalment} / month"))
    if rate:
        rows.append(("**Rate of interest**", f"{rate} (floating)"))
    if apr:
        rows.append(("**APR**", apr))
    if subvention:
        rows.append(("**Subvention**", subvention))
    if repayment:
        rows.append(("**Repayment mode**", repayment))
    if loan_status and disbursement_type:
        rows.append(
            ("**Loan status**", f"{loan_status} ({disbursement_type.lower()})")
        )
    elif loan_status:
        rows.append(("**Loan status**", loan_status))
    rows.append(("**Branch**", "Gurgaon"))
    rows.append(
        (
            "**Unit / property**",
            "3051, Tower 3, ATS Grandstand Phase 1, Sector 99A",
        )
    )
    rows.append(
        (
            "**Disbursement to**",
            "ATS Realworth Pvt Ltd Escrow A/c, Yes Bank (20-Nov-2018, Chq 217689)",
        )
    )
    if principal_paid:
        rows.append(("**Total principal paid**", f"Rs. {principal_paid}"))
    if interest_paid:
        rows.append(("**Total interest paid**", f"Rs. {interest_paid}"))
    if charges_due:
        rows.append(("**Charges/fee due (as on statement)**", f"Rs. {charges_due}"))
    if excess is not None:
        rows.append(("**Excess amount (as on statement)**", f"Rs. {excess}"))
    if balance_tenure:
        tenure = (
            balance_tenure
            if "month" in balance_tenure.lower()
            else f"{balance_tenure} months"
        )
        rows.append(("**Balance loan tenure**", tenure))

    return rows


def main() -> None:
    doc = fitz.open(PDF)
    page_texts: list[str] = []
    text_pages = 0

    for page in doc:
        text = page.get_text().strip()
        page_texts.append(text)
        if text:
            text_pages += 1

    page_count = doc.page_count
    doc.close()

    key_rows = extract_key_details(page_texts[0])
    subvention_value = next(
        (v for k, v in key_rows if "Subvention" in k),
        None,
    )

    parts = [
        "# Statement of account (first page) — PNB Housing\n\n",
        f"Converted from: `{SOURCE_NAME}`\n\n",
        "**Note:** First page of Statement of Account dated **22.06.2026** "
        "(period 01/04/2019 to 22/06/2026; summary and disbursement details only).\n\n",
        f"**Pages:** {page_count}\n\n---\n",
    ]

    for i, text in enumerate(page_texts, start=1):
        parts.append(f"\n## Page {i}\n\n")
        if text:
            parts.append("```text\n" + text + "\n```\n")
        else:
            parts.append("*(No extractable text — image-only/scanned page.)*\n")

    parts.append("\n---\n\n## Key details (extracted)\n\n")
    parts.append("| Field | Value |\n|-------|--------|\n")
    for field, value in key_rows:
        parts.append(f"| {field} | {value} |\n")

    parts.append("\n## Extraction note\n\n")
    parts.append(f"- **Source PDF:** `attachments/{SOURCE_NAME}`\n")
    parts.append(f"- **Total pages:** {page_count}\n")
    parts.append(f"- **Pages with extractable text:** {text_pages}\n")
    if text_pages < page_count:
        parts.append(
            "- **Scanned/image-only pages:** "
            f"{page_count - text_pages} page(s) had no extractable text.\n"
        )
    else:
        parts.append(
            "- **Document type:** Native/digital PDF with extractable text "
            "(not a scanned image).\n"
        )
    if subvention_value:
        parts.append(
            f"- **Subvention:** **{subvention_value}** — subvention field appears "
            "on the statement summary.\n"
        )
    else:
        parts.append(
            "- **Subvention:** Not found in extracted text on this page.\n"
        )

    OUT.write_text("".join(parts), encoding="utf-8")
    print(f"Wrote {OUT} ({OUT.stat().st_size} bytes, {page_count} pages)")


if __name__ == "__main__":
    main()
