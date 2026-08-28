"""Convert 25062026*.pdf files to markdown in suitable folders."""
from __future__ import annotations

from pathlib import Path

import fitz

REPO = Path(__file__).resolve().parents[5]
DTCP_BASE = Path(__file__).resolve().parent.parent
DTCP_MD = Path(__file__).resolve().parent
PNB_MD = REPO / "PNB_ATS" / "email_letters" / "md"

JOBS: list[dict] = [
    {
        "pdf": DTCP_BASE / "25062026_Affidavit_cum_undertaking.pdf",
        "out": DTCP_MD / "25062026_Affidavit_cum_undertaking.md",
        "title": "25062026 — Affidavit-cum-undertaking (ATS / builder)",
        "note": (
            "Scanned affidavit-cum-undertaking submitted by ATS/builder group dated **25.06.2026**, "
            "enclosed with allottee email to PNBHFL (`krishan.kumar@pnbhousing.com`). "
            "Relates to Unit **3051**, ATS Grandstand Phase-I — pre-EMI / HARERA compliance undertaking."
        ),
        "rel_source": "../25062026_Affidavit_cum_undertaking.pdf",
    },
    {
        "pdf": DTCP_BASE / "25062026_ATS_Signed _Ledger.pdf",
        "out": DTCP_MD / "25062026_ATS_Signed_Ledger.md",
        "title": "25062026 — ATS signed payment ledger",
        "note": (
            "Scanned **signed ledger** from ATS/builder dated **25.06.2026** — payment/account statement "
            "for Unit **3051** pre-EMI / dues. Enclosed with allottee email to PNBHFL."
        ),
        "rel_source": "../25062026_ATS_Signed _Ledger.pdf",
    },
    {
        "pdf": DTCP_BASE / "25062026_Consent letter.pdf",
        "out": DTCP_MD / "25062026_Consent_letter.md",
        "title": "25062026 — Consent letter (revised building plan)",
        "note": (
            "Scanned **consent letter** dated **25.06.2026** from allottee regarding revised building plan "
            "(FAR, density, parking, green area) — conditional on HARERA compliance and compensation rights."
        ),
        "rel_source": "../25062026_Consent letter.pdf",
    },
    {
        "pdf": PNB_MD.parent
        / "pdf"
        / "Gmail - 25062026 Subject_ Request for direct pre-EMI payment by builder_promoter into loan account , Loan A_c HOU_GUR_0918_578047.pdf",
        "out": PNB_MD / "25062026_Gmail_REQUEST_Direct_PreEMI_Payment_Builder_Unit_3051.md",
        "title": "25062026 — Gmail to PNBHFL — direct pre-EMI payment by builder",
        "note": (
            "Gmail export: allottee email dated **25.06.2026 at 16:01** to **krishan.kumar@pnbhousing.com** "
            "requesting direct pre-EMI payment by builder/promoter into Loan A/c **HOU/GUR/0918/578047** "
            "(Unit **3051**). **7 attachments** including affidavit, statement, HARERA orders, builder pre-EMI letter."
        ),
        "rel_source": "../pdf/Gmail - 25062026 Subject_ Request for direct pre-EMI payment by builder_promoter into loan account , Loan A_c HOU_GUR_0918_578047.pdf",
        "key_table": [
            ("**From**", "Vijyant Agarwal <prof.vijyant.agarwal@gmail.com>"),
            ("**To**", "krishan.kumar@pnbhousing.com"),
            ("**Email date**", "25 June 2026 at 16:01"),
            (
                "**Subject**",
                "Request for direct pre-EMI payment by builder/promoter into loan account, "
                "Loan A/c HOU/GUR/0918/578047",
            ),
            ("**Customer ID**", "1556854"),
            ("**Loan Account No.**", "HOU/GUR/0918/578047"),
            ("**Subvention**", "Yes (per statement dated 22.06.2026)"),
            ("**Attachments (7)**", "FIR notice; affidavit; statement; Gmail request; pre-EMI letter; HARERA order 08.08.2024; RERA order 18.03.2025"),
        ],
    },
]


def extract_pages(pdf_path: Path) -> list[str]:
    doc = fitz.open(pdf_path)
    pages = [page.get_text().strip() for page in doc]
    doc.close()
    return pages


def build_md(job: dict, pages: list[str]) -> str:
    page_count = len(pages)
    text_pages = sum(1 for p in pages if p)

    parts = [
        f"# {job['title']}\n\n",
        f"Converted from: `{job['pdf'].name}`\n\n",
        f"**Source:** `{job['rel_source']}`\n\n",
        f"**Note:** {job['note']}\n\n",
        f"**Pages:** {page_count} ({text_pages} with extractable text)\n\n",
        "---\n\n",
    ]

    for i, text in enumerate(pages, start=1):
        parts.append(f"## Page {i}\n\n")
        if text:
            parts.append(f"```text\n{text}\n```\n\n")
        else:
            parts.append("*(No extractable text — image/scanned page.)*\n\n")

    key_table: list[tuple[str, str]] | None = job.get("key_table")
    if key_table:
        parts.append("---\n\n## Key details (extracted)\n\n")
        parts.append("| Field | Value |\n|-------|--------|\n")
        for field, value in key_table:
            parts.append(f"| {field} | {value} |\n")
        parts.append("\n")

    if text_pages < page_count:
        parts.append(
            "## Extraction note\n\n"
            f"- **Scanned/image-only pages:** {page_count - text_pages} of {page_count} "
            "had no extractable text layer. Refer to source PDF for full content.\n"
        )

    return "".join(parts)


def main() -> None:
    for job in JOBS:
        pdf: Path = job["pdf"]
        out: Path = job["out"]
        if not pdf.is_file():
            print(f"Skip (missing): {pdf}")
            continue
        pages = extract_pages(pdf)
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(build_md(job, pages), encoding="utf-8")
        print(f"Saved: {out}")


if __name__ == "__main__":
    main()
