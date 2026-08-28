"""Convert 105_ATS_Builder_preEMI_Letter_23102018.pdf to markdown."""
from pathlib import Path

import fitz

PNB_ROOT = Path(__file__).resolve().parents[1]
ATT = PNB_ROOT / "attachments"
PDF = ATT / "105_ATS_Builder_preEMI_Letter_23102018.pdf"
OUT = Path(__file__).resolve().parent / "105_ATS_Builder_preEMI_Letter_23102018.md"


def resolve_pdf() -> Path:
    if PDF.exists():
        return PDF
    raise FileNotFoundError(f"No subvention letter PDF found ({PDF.name})")

# Single-page scanned letter; no pymupdf text layer. Transcription from source image.
SCANNED_TRANSCRIPTION = """\
Date: 23-10-2018

To,
Mr. Vijyant Agarwal S/o Mr. Braham Prakash Agarwal
Mrs. Vineeta Agarwal W/o Mr. Vijyant Agarwal
R/o House No.-07, Type-05, Netaji Subhas Institute of Technology,
Sector-03, Dwarka, New Delhi-110078.

Ref.: Agreement for Sale dated 18-10-2018/Booking Reg. No. 27

Subject: Regarding payment of Pre EMI till offer of Possession is made by the
Developer to Buyer with respect to Apartment no. 3051, in 'ATS Grandstand
Phase-I' situated at Sec 99A, Gurgaon.

Dear Sir/Madam,

This has reference to your booking/application for purchase of an apartment
in our project ATS Grandstand Phase-I situated at Sector 99A, Gurgaon and
your request for payment of Pre-EMI by us till the offer of possession is made
by the Developer to the Buyer.

We hereby confirm that after expiry of subvention scheme (approved by PNBHFL
BANK), we shall pay the Pre-EMI every month to you or directly to bank, as the
case may be, till such time the offer of Possession is made by the Developer
to the Buyer in terms of clause 7.1 of the Agreement for Sale.

Once the possession notice is issued by the Developer, this scheme shall
stand closed and the Buyer shall be solely liable to pay all the remaining
EMIs on the bank loan.

The Buyer undertakes to take possession of the apartment within 30 days of
receipt of possession notice, after completing the formalities mentioned in
clause 7.2 of the Agreement for Sale.

It is further agreed between the parties that the Buyer shall not be entitled
to any delay compensation or any other compensation as provided in clause 7.6
of the Agreement for Sale dated 18-10-2018.

All other terms and conditions of the Agreement for Sale dated 18-10-2018
shall remain unchanged. This letter shall form part of the Agreement for Sale.

For M/s ATS Realworth Private Limited

[Signed]
Anand
Authorized Signatory

ATS Realworth Pvt. Ltd.
Registered Office: 711/92, Deepali, Nehru Place, New Delhi - 110019
Corporate Office: ATS Tower, Plot No.16, Sector 135, Noida - 201305
"""


def main() -> None:
    source_pdf = resolve_pdf()
    source_name = source_pdf.name
    doc = fitz.open(source_pdf)
    page_count = doc.page_count
    text_pages = 0
    image_only_pages = 0

    parts = [
        "# ATS Subvention letter — 2018\n\n",
        f"Converted from: `{source_name}`\n\n",
        "**Note:** Builder letter dated **23.10.2018** confirming Pre-EMI payment "
        "by ATS Realworth Pvt. Ltd. until offer of possession (subvention scheme "
        "approved by PNBHFL).\n\n",
        f"**Pages:** {page_count}\n\n---\n",
    ]

    for i, page in enumerate(doc, start=1):
        text = page.get_text().strip()
        has_images = bool(page.get_images())
        parts.append(f"\n## Page {i}\n\n")

        if text:
            text_pages += 1
            parts.append("```text\n" + text + "\n```\n")
        elif has_images:
            image_only_pages += 1
            parts.append(
                "*(Scanned/image-only page — no extractable text layer in PDF. "
                "Transcription below from visual review of source.)*\n\n"
            )
            parts.append("```text\n" + SCANNED_TRANSCRIPTION.strip() + "\n```\n")
        else:
            parts.append("*(No extractable text.)*\n")

    doc.close()

    parts.append(
        "\n---\n\n## Key details (extracted)\n\n"
        "| Field | Value |\n"
        "|-------|--------|\n"
        "| **Document date** | 23-10-2018 |\n"
        "| **Developer** | M/s ATS Realworth Private Limited |\n"
        "| **Buyers** | Mr. Vijyant Agarwal & Mrs. Vineeta Agarwal |\n"
        "| **Unit** | Apartment No. 3051 |\n"
        "| **Project** | ATS Grandstand Phase-I, Sector 99A, Gurgaon |\n"
        "| **Agreement for Sale** | Dated 18-10-2018; Booking Reg. No. 27 |\n"
        "| **Financing bank** | PNBHFL (PNB Housing Finance Limited) |\n"
        "| **Subvention / Pre-EMI period** | After expiry of bank-approved subvention scheme, "
        "until offer of possession under clause 7.1 of Agreement for Sale |\n"
        "| **Developer obligation** | Pay Pre-EMI monthly to buyer or directly to bank |\n"
        "| **End of scheme** | On issuance of possession notice; buyer then liable for all remaining EMIs |\n"
        "| **Buyer obligation** | Take possession within 30 days of possession notice (clause 7.2) |\n"
        "| **Waiver** | Buyer not entitled to delay/other compensation under clause 7.6 |\n"
        "| **Signatory** | Anand, Authorized Signatory |\n"
        "\n## Extraction note\n\n"
        f"- **Source PDF:** `attachments/{source_name}`\n"
        f"- **Total pages:** {page_count}\n"
        f"- **Pages with extractable text:** {text_pages}\n"
        f"- **Scanned/image-only pages:** {image_only_pages} "
        "(embedded JPEG; pymupdf returned no text layer)\n"
        "- **Transcription:** Page 1 content transcribed manually from rendered page "
        "because the source PDF is image-only.\n"
    )

    OUT.write_text("".join(parts), encoding="utf-8")
    print(f"Wrote {OUT} ({OUT.stat().st_size} bytes, {page_count} pages)")


if __name__ == "__main__":
    main()
