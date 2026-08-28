"""Convert 24062026 ATS Gmail reply PDFs to markdown in md_files/."""
from __future__ import annotations

import re
from pathlib import Path

import fitz

BASE = Path(__file__).resolve().parent.parent
OUT = Path(__file__).resolve().parent

PDFS = [
    BASE / "24062026_gmail_reply_by_ATS_with_attachment.PDF",
    BASE / "24062026_gmail_atachment_reply_by_ATS.pdf",
]


def extract_pages(pdf_path: Path) -> list[str]:
    doc = fitz.open(pdf_path)
    pages = [page.get_text().strip() for page in doc]
    doc.close()
    return pages


def build_md(pdf_path: Path, pages: list[str]) -> str:
    page_count = len(pages)
    text_pages = sum(1 for p in pages if p)
    rel = f"../{pdf_path.name}"

    if "with_attachment" in pdf_path.name.lower():
        title = "24062026 — Gmail reply from ATS (with attachment)"
        note = (
            "Gmail export: ATS/Homekraft reply dated **24.06.2026** to allottee email "
            "dated **22.06.2026** on STP meeting / revised building plan objections. "
            "Attachment: colonizer reply table (`24062026_gmail_atachment_reply_by_ATS.pdf`)."
        )
    else:
        title = "24062026 — ATS attachment — reply to allottee objections (STP / revised plan)"
        note = (
            "Attachment to Gmail dated **24.06.2026** from `grandstand.customercare@homekraft.in`. "
            "Colonizer point-wise reply to objections raised by **Unit 3051** (email **22.06.2026**)."
        )

    parts = [
        f"# {title}\n\n",
        f"Converted from: `{pdf_path.name}`\n\n",
        f"**Source:** `{rel}`\n\n",
        f"**Note:** {note}\n\n",
        f"**Pages:** {page_count} ({text_pages} with extractable text)\n\n",
        "---\n\n",
    ]

    for i, text in enumerate(pages, start=1):
        parts.append(f"## Page {i}\n\n")
        if text:
            parts.append(f"```text\n{text}\n```\n\n")
        else:
            parts.append("*(No extractable text — likely image/scanned.)*\n\n")

    return "".join(parts)


def main() -> None:
    for pdf in PDFS:
        if not pdf.exists():
            print(f"Skip (missing): {pdf}")
            continue
        pages = extract_pages(pdf)
        md_path = OUT / f"{pdf.stem}.md"
        md_path.write_text(build_md(pdf, pages), encoding="utf-8")
        print(f"Saved: {md_path}")


if __name__ == "__main__":
    main()
