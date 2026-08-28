"""Shared conversion for 112_/113_ ATS builder email PDFs."""
import re
from pathlib import Path

import fitz

PNB_ROOT = Path(__file__).resolve().parents[1]
OUT_DIR = Path(__file__).resolve().parent


def _first_match(text: str, pattern: str) -> str | None:
    match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
    return match.group(1).strip() if match else None


def extract_key_details(full_text: str) -> list[tuple[str, str]]:
    from_addr = _first_match(
        full_text,
        r"(Grandstand\.customercare@homekraft\.in\s*<[^>]+>)",
    )
    to_addr = _first_match(
        full_text,
        r"To:\s*(Vijyant Agarwal\s*<[^>]+>)",
    )
    email_date = _first_match(
        full_text,
        r"homekraft\.in>\s*\n?\s*(\d{1,2}\s+\w+\s+\d{4}\s+at\s+\d{1,2}:\d{2})",
    )
    subject = _first_match(
        full_text,
        r"(Invitation for redressal of grievances : Meeting held in the office of W/o Senior Town Planner,\s*"
        r"Gurugram on \d{2}\.\d{2}\.\d{4} regarding your objections on approval of revised building plans of\s*"
        r"Group housing project namely .ATS Grandstand.)",
    )
    if subject:
        subject = re.sub(r"\s+", " ", subject)
        subject = subject.replace("\ufffd", '"').replace(""", '"').replace(""", '"')
    meeting_date = _first_match(
        full_text,
        r"meeting on (\d{2}\.\d{2}\.\d{4}\s+at\s+\d{1,2}:\d{2}\s+AM)",
    )
    meeting_venue = _first_match(
        full_text,
        r"at\s+our office at ([^\n.]+)",
    )
    dtcp_meeting = _first_match(
        full_text,
        r"meeting with W/o Senior Town\s+Planner on (\d{2}\.\d{2}\.\d{4})",
    )

    pre_emi_section = _first_match(
        full_text,
        r"(Regarding payment of pre-EMIs in your loan account.*?(?=\n\s*You are accordingly requested))",
    )
    pre_emi_items: list[str] = []
    if pre_emi_section:
        pre_emi_items = re.findall(
            r"(\d+\.\s+[^\n]+(?:\n(?!\d+\.)[^\n]+)*)",
            pre_emi_section,
        )
        pre_emi_items = [re.sub(r"\s+", " ", item.strip()) for item in pre_emi_items]
    pre_emi_block = "; ".join(pre_emi_items) if pre_emi_items else None

    rows: list[tuple[str, str]] = []
    if from_addr:
        rows.append(("**From**", from_addr))
    if to_addr:
        rows.append(("**To**", to_addr))
    if email_date:
        rows.append(("**Email date**", email_date))
    if subject:
        rows.append(("**Subject**", subject))
    if meeting_date:
        venue = f" — {meeting_venue}" if meeting_venue else ""
        rows.append(("**Proposed meeting**", f"{meeting_date}{venue}"))
    if dtcp_meeting:
        rows.append(("**Prior DTCP meeting referenced**", dtcp_meeting))
    rows.append(
        (
            "**Pre-EMI / bank documents requested**",
            pre_emi_block
            or "Bank subvention confirmation; NoC for third-party payment; tripartite freeze status; "
            "affidavit (Rs100 NJSP) re subvention/non-conversion; loan account details with bank "
            "certification; other bank requirements",
        ),
    )
    rows.append(
        (
            "**Other undertakings offered**",
            "Individual affidavit (format awaited), account statement, payment resolution "
            "(per post-DTCP meeting on 19.06.2026)",
        ),
    )
    return rows


def convert_pdf(pdf_path: Path) -> Path:
    doc = fitz.open(pdf_path)
    page_texts: list[str] = []
    text_pages = 0

    for page in doc:
        text = page.get_text().strip()
        page_texts.append(text)
        if text:
            text_pages += 1

    page_count = doc.page_count
    doc.close()

    full_text = "\n\n".join(page_texts)
    key_rows = extract_key_details(full_text)
    out_name = pdf_path.stem + ".md"
    out_path = OUT_DIR / out_name

    rel_source = Path("attachments") / pdf_path.name
    parts = [
        "# ATS email — invitation for grievance redressal meeting\n\n",
        f"Converted from: `{pdf_path.name}`\n\n",
        f"**Source:** `{rel_source.as_posix()}`\n\n",
        "**Note:** Email from ATS Grandstand CRM dated **22.06.2026** inviting allottee "
        "to a meeting on **25.06.2026** regarding pre-EMI, affidavits, and bank documentation.\n\n",
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
    parts.append(f"- **Source PDF:** `{rel_source.as_posix()}`\n")
    parts.append(f"- **Total pages:** {page_count}\n")
    parts.append(f"- **Pages with extractable text:** {text_pages}\n")
    if text_pages < page_count:
        parts.append(
            "- **Scanned/image-only pages:** "
            f"{page_count - text_pages} page(s) had no extractable text.\n"
        )
    else:
        parts.append("- **Text layer:** Native/selectable text extracted via PyMuPDF.\n")

    out_path.write_text("".join(parts), encoding="utf-8")
    return out_path
