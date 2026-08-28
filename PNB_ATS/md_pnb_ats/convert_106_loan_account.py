"""Convert 106_PNB_Loan_Account_Details.pdf to markdown."""
from pathlib import Path

import fitz

REF = Path(__file__).resolve().parents[1] / "attachments"
PDF = REF / "106_PNB_Loan_Account_Details.pdf"
OUT = Path(__file__).resolve().parent / "106_PNB_Loan_Account_Details.md"


def main() -> None:
    doc = fitz.open(PDF)
    text_pages = 0
    parts = [
        "# Loan account details — PNB Housing\n\n",
        "Converted from: `106_PNB_Loan_Account_Details.pdf`\n\n",
        "**Note:** Statement of Account dated **05.11.2021** (sanction/disbursement summary).\n\n",
        f"**Pages:** {doc.page_count}\n\n---\n",
    ]
    for i, page in enumerate(doc, start=1):
        text = page.get_text().strip()
        parts.append(f"\n## Page {i}\n\n")
        if text:
            text_pages += 1
            parts.append("```text\n" + text + "\n```\n")
        else:
            parts.append("*(No extractable text.)*\n")
    page_count = doc.page_count
    doc.close()

    parts.append(
        "\n---\n\n## Key details (extracted)\n\n"
        "| Field | Value |\n"
        "|-------|--------|\n"
        "| **Customer ID** | 1556854 |\n"
        "| **Loan Account No.** | HOU/GUR/0918/578047 |\n"
        "| **Statement date** | 05-Nov-2021 |\n"
        "| **Sanction date** | 19-Sep-2018 |\n"
        "| **Sanction amount** | Rs. 66,45,815.00 |\n"
        "| **Disbursed (as on statement)** | Rs. 26,59,789.00 |\n"
        "| **Current instalment** | Rs. 20,725.00 / month |\n"
        "| **Rate of interest** | 9.35% p.a. (floating) |\n"
        "| **Subvention** | Yes |\n"
        "| **Repayment mode** | ECS |\n"
        "| **Loan status** | Active (partly disbursed) |\n"
        "| **Branch** | Gurgaon |\n"
        "| **Unit / property** | 3051, Tower 3, ATS Grandstand Phase 1, Sector 99A |\n"
        "| **Disbursement to** | ATS Realworth Pvt Ltd Escrow A/c, Yes Bank (20-Nov-2018, Chq 217689) |\n"
        "| **Excess amount (as on statement)** | Rs. 2,907.68 |\n"
        "| **Closing balance (05/11/2021)** | -2,322.98 (credit) |\n"
        "\n## Extraction note\n\n"
        f"- **Source PDF:** `attachments/106_PNB_Loan_Account_Details.pdf`\n"
        f"- **Total pages:** {page_count}\n"
        f"- **Pages with extractable text:** {text_pages}\n"
    )
    OUT.write_text("".join(parts), encoding="utf-8")
    print(f"Wrote {OUT} ({OUT.stat().st_size} bytes)")


if __name__ == "__main__":
    main()
