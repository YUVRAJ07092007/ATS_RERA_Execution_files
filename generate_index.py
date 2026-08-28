"""Generate file indexes for ATS_RERA_execution folder."""
from pathlib import Path
from collections import defaultdict
from datetime import datetime
import hashlib
import re

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

root = Path(__file__).parent
out_md = root / "00_MASTER_FILE_INDEX.md"
out_docx = root / "00_MASTER_FILE_INDEX.docx"
main_md = root / "01_MAIN_FILES_INDEX.md"
main_docx = root / "01_MAIN_FILES_INDEX.docx"
dispatch_md = root / "02_DISPATCH_CORRESPONDENCE_INDEX.md"
dispatch_docx = root / "02_DISPATCH_CORRESPONDENCE_INDEX.docx"
skip = {".git", "__pycache__"}

SKIP_PATH_PARTS = {
    "md_files",
    "md_pnb_ats",
    "md_files_reply_by_ATS",
    "md_files_RERA_Approval_August_2026",
    "__pycache__",
}
MAIN_SKIP_EXT = {".md", ".py", ".gdoc", ".gsheet", ".sample"}
MAIN_SKIP_NAMES = {
    "README.md",
    "desktop.ini",
    ".gitignore",
    "generate_index.py",
    "Ats execution pdf 241024.pdf",
}
INDEX_OUTPUT_PREFIXES = (
    "00_MASTER_FILE_INDEX",
    "01_MAIN_FILES_INDEX",
    "02_DISPATCH_CORRESPONDENCE_INDEX",
)


def should_skip(p: Path) -> bool:
    return bool(set(p.parts) & skip)


def is_index_output(name: str) -> bool:
    return name.startswith(INDEX_OUTPUT_PREFIXES) or name in MAIN_SKIP_NAMES


def is_main_file(f: Path) -> bool:
    if not f.is_file() or should_skip(f):
        return False
    if any(part in SKIP_PATH_PARTS for part in f.parts):
        return False
    if is_index_output(f.name):
        return False
    if f.suffix.lower() in MAIN_SKIP_EXT:
        return False
    return True


def file_content_hash(path: Path) -> str:
    digest = hashlib.md5()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def canonical_path_score(rel: Path) -> tuple:
    """Lower score = preferred canonical location when deduplicating."""
    text = str(rel).lower()
    score = 0
    if "02062026_annexures" in text or "\\annexure_" in text.replace("/", "\\"):
        score += 100
    if "annexures" in text and "mom" in text:
        score += 40
    if "dtpc_hrera_conditional_approval" in text and rel.name[0].isdigit():
        score += 30
    if "fabrication" in text and text.count("fabrication") > 1:
        score += 20
    if " (1)" in rel.name.lower():
        score += 15
    if "pdfs" in text:
        score -= 10
    depth = len(rel.parts)
    if depth <= 2:
        score -= 15
    return (score, depth, str(rel).lower())


def collect_unique_main_files() -> tuple[dict[str, list[str]], int, int]:
    """Return unique main files by folder, plus raw and skipped duplicate counts."""
    by_hash: dict[str, list[Path]] = defaultdict(list)

    for f in sorted(root.rglob("*")):
        if not is_main_file(f):
            continue
        by_hash[file_content_hash(f)].append(f.relative_to(root))

    files_by_dir: dict[str, list[str]] = defaultdict(list)
    raw_count = sum(len(paths) for paths in by_hash.values())
    duplicate_copies = 0

    for paths in by_hash.values():
        canonical = sorted(paths, key=canonical_path_score)[0]
        duplicate_copies += len(paths) - 1
        parent = (
            str(canonical.parent).replace("\\", "/")
            if str(canonical.parent) != "."
            else "(ROOT)"
        )
        files_by_dir[parent].append(canonical.name)

    for parent in files_by_dir:
        files_by_dir[parent].sort(key=str.lower)

    return files_by_dir, raw_count, duplicate_copies


def categorize_root(name: str) -> str:
    n = name.lower()
    if any(
        x in n
        for x in [
            "execution",
            "speedpost",
            "receipt",
            "vakalatnama",
            "demand draft",
            "check list",
            "haryanarera",
            "index",
            "brief facts",
        ]
    ):
        return "Execution Filing (Oct 2024)"
    if "affidavit" in n:
        return "Affidavits"
    if "outstanding" in n or "pre emi" in n or "payment detail" in n:
        return "Outstanding Dues & Pre-EMI"
    if "cheque" in n:
        return "Payments by ATS (root copies)"
    if "pnb" in n or ("statement" in n and "ats" not in n):
        return "Bank / PNB Statements"
    if "order" in n or "rera" in n:
        return "HARERA Orders"
    if n.endswith(".gdoc") or n.endswith(".gsheet") or "untitled" in n:
        return "Drafts & Working Files"
    if "brochure" in n or "regulation" in n:
        return "Other / Reference"
    return "Other / Reference"


SECTIONS = [
    ("(ROOT)", "1. Root Folder — Original Execution & Core Documents", None),
    ("Payment by ATS", "2. Payment by ATS", "Payments made by builder ATS to allottee"),
    ("DTTP complaint", "3. DTCP / DTTP Complaint & RTI", "Complaints and RTI to DTCP and HARERA"),
    (
        "PNB_ATS",
        "4. PNB_ATS — Pre-EMI / Bank Correspondence",
        "See also PNB_ATS/README.md for detailed 101–114 index",
    ),
    ("PNB_ATS/attachments", "4a. PNB_ATS — Attachments (PDFs 101–114)", None),
    ("PNB_ATS/email_letters/md", "4b. PNB_ATS — Email/Letter Drafts (Markdown)", None),
    ("PNB_ATS/email_letters/docx", "4c. PNB_ATS — Email/Letter Drafts (Word)", None),
    ("PNB_ATS/email_letters/pdf", "4d. PNB_ATS — Sent Email Exports (PDF)", None),
    ("PNB_ATS/md_pnb_ats", "4e. PNB_ATS — Markdown Conversions & Scripts", None),
    (
        "Change_of_devloper_by_ATS",
        "5. Change of Developer by ATS",
        "ATS attempt to change developer — background docs",
    ),
    ("Change_of_devloper_by_ATS/pdfs", "5a. Change of Developer — Source PDFs", None),
    ("Change_of_devloper_by_ATS/Letter to all", "5b. Change of Developer — Letters to Allottees", None),
    (
        "Change_of_devloper_by_ATS/Reply by ATS",
        "5c. ATS Reply Documents (Original PDFs)",
        "Builder ATS submissions numbered 0–8",
    ),
    ("Change_of_devloper_by_ATS/Reply by ATS/md_files", "5d. ATS Reply — Markdown Conversions", None),
    (
        "Change_of_devloper_by_ATS/Reply by ATS/reply by Allotee",
        "5e. Allottee Reply to ATS (01.06.2026 filing)",
        "Main application + rejoinder",
    ),
    (
        "Change_of_devloper_by_ATS/Reply by ATS/reply by Allotee/02062026_Annexures",
        "5f. Filing Annexures Index",
        "See 00_ANNEXURE_INDEX.md in this folder",
    ),
    (
        "Change_of_devloper_by_ATS/Reply by ATS/reply by Allotee/02062026_Annexures/Annexure_1_HARERA_Orders",
        "5f-1. Annexure 1 — HARERA Orders",
        None,
    ),
    (
        "Change_of_devloper_by_ATS/Reply by ATS/reply by Allotee/02062026_Annexures/Annexure_2_Execution_Proceedings",
        "5f-2. Annexure 2 — Execution Proceedings",
        None,
    ),
    (
        "Change_of_devloper_by_ATS/Reply by ATS/reply by Allotee/02062026_Annexures/Annexure_3_Allottee_Representations",
        "5f-3. Annexure 3 — Allottee Representations",
        None,
    ),
    (
        "Change_of_devloper_by_ATS/Reply by ATS/reply by Allotee/02062026_Annexures/Annexure_4_ATS_Reply_and_Revised_Plans",
        "5f-4. Annexure 4 — ATS Reply & Revised Plans",
        None,
    ),
    (
        "Change_of_devloper_by_ATS/Reply by ATS/reply by Allotee/02062026_Annexures/Annexure_5_Change_of_Developer_Background",
        "5f-5. Annexure 5 — Change of Developer Background",
        None,
    ),
    (
        "Change_of_devloper_by_ATS/Reply by ATS/reply by Allotee/DTPC_HRERA_CONDITIONAL_APPROVAL",
        "5g. DTCP / HARERA Conditional Approval",
        "Conditional approval application and hearing docs",
    ),
    (
        "Change_of_devloper_by_ATS/Reply by ATS/reply by Allotee/DTPC_HRERA_CONDITIONAL_APPROVAL/md_files",
        "5g-1. DTCP/HARERA — Markdown Files",
        None,
    ),
    (
        "Change_of_devloper_by_ATS/Reply by ATS/reply by Allotee/DTPC_HRERA_CONDITIONAL_APPROVAL/MOM_dated_18062026_DTCP",
        "5g-2. DTCP Hearing MOM (18.06.2026)",
        None,
    ),
    (
        "Change_of_devloper_by_ATS/Reply by ATS/reply by Allotee/DTPC_HRERA_CONDITIONAL_APPROVAL/MOM_dated_18062026_DTCP/Annexures",
        "5g-3. DTCP Hearing — Annexures",
        None,
    ),
    (
        "Change_of_devloper_by_ATS/Reply by ATS/reply by Allotee/DTPC_HRERA_CONDITIONAL_APPROVAL/MOM_dated_18062026_DTCP/Annexures/6.3_Email_30052026_Builder",
        "5g-4. DTCP Hearing — Builder Email (30.05.2026)",
        None,
    ),
    (
        "Change_of_devloper_by_ATS/Reply by ATS/reply by Allotee/Fabrication",
        "5h. Fabrication Evidence",
        "Evidence of fabricated ATS documents",
    ),
    (
        "Change_of_devloper_by_ATS/Reply by ATS/reply by Allotee/md_files_reply_by_ATS",
        "5i. Allottee Reply — ATS Docs (Markdown)",
        None,
    ),
    (
        "Change_of_devloper_by_ATS/Reply by ATS/md_files_reply_by_ATS",
        "5j. ATS Reply — Additional Markdown",
        None,
    ),
    (
        "Change_of_devloper_by_ATS/approval_by_RERA_August_2026",
        "5k. HARERA Approval — Continuation of Registration (August 2026)",
        "RERA-GRG-2264-2026; hearing 10.08.2026; RC No. 06 of 2018 — Phase-I continuation u/s 7(3)",
    ),
    (
        "Change_of_devloper_by_ATS/approval_by_RERA_August_2026/md_files_RERA_Approval_August_2026",
        "5k-1. HARERA Approval — Markdown Transcription",
        None,
    ),
]

MAIN_SECTIONS = [
    ("(ROOT)", "1. Root Folder — Execution & Core Documents", None),
    ("Payment by ATS", "2. Payment by ATS", "Payments made by builder ATS to allottee"),
    ("DTTP complaint", "3. DTCP / DTTP Complaint & RTI", "Complaints and RTI to DTCP and HARERA"),
    ("PNB_ATS/attachments", "4. PNB — Supporting PDFs (101–114)", "Canonical PDF attachments for PNB correspondence"),
    ("PNB_ATS/email_letters/docx", "5. PNB — Outgoing Letters (Word)", None),
    ("PNB_ATS/email_letters/pdf", "6. PNB — Sent Emails (PDF)", None),
    (
        "Change_of_devloper_by_ATS",
        "7. Change of Developer — Background",
        "ATS change-of-developer background documents",
    ),
    ("Change_of_devloper_by_ATS/pdfs", "7a. Change of Developer — Source PDFs", None),
    ("Change_of_devloper_by_ATS/Letter to all", "7b. Letters to Allottees", None),
    (
        "Change_of_devloper_by_ATS/Reply by ATS",
        "7c. ATS Reply (Original Documents)",
        "Builder ATS submissions numbered 0–8",
    ),
    (
        "Change_of_devloper_by_ATS/Reply by ATS/reply by Allotee",
        "7d. Allottee Filings & Replies",
        "Applications, rejoinders and representations",
    ),
    (
        "Change_of_devloper_by_ATS/Reply by ATS/reply by Allotee/DTPC_HRERA_CONDITIONAL_APPROVAL",
        "7e. DTCP / HARERA Conditional Approval",
        "Conditional approval application and hearing documents",
    ),
    (
        "Change_of_devloper_by_ATS/Reply by ATS/reply by Allotee/DTPC_HRERA_CONDITIONAL_APPROVAL/MOM_dated_18062026_DTCP",
        "7f. DTCP Hearing MOM (18.06.2026)",
        None,
    ),
    (
        "Change_of_devloper_by_ATS/Reply by ATS/reply by Allotee/DTPC_HRERA_CONDITIONAL_APPROVAL/MOM_dated_18062026_DTCP/Annexures",
        "7g. DTCP Hearing — Annexures",
        None,
    ),
    (
        "Change_of_devloper_by_ATS/Reply by ATS/reply by Allotee/Fabrication",
        "7h. Fabrication Evidence",
        "Evidence of fabricated ATS documents",
    ),
    (
        "Change_of_devloper_by_ATS/approval_by_RERA_August_2026",
        "7i. HARERA Approval — Continuation of Registration (August 2026)",
        "RERA-GRG-2264-2026; hearing 10.08.2026; RC No. 06 of 2018 — Phase-I u/s 7(3)",
    ),
]

ROOT_CAT_ORDER = [
    "Execution Filing (Oct 2024)",
    "HARERA Orders",
    "Affidavits",
    "Bank / PNB Statements",
    "Outstanding Dues & Pre-EMI",
    "Payments by ATS (root copies)",
    "Other / Reference",
    "Drafts & Working Files",
]


def _strip_md_inline(text: str) -> str:
    """Remove backticks for plain table cell text."""
    return text.strip().strip("`").strip()


def _parse_table_cells(line: str) -> list[str]:
    """Parse a markdown table row; drop empty edge cells from pipe formatting."""
    cells = [_strip_md_inline(c) for c in line.strip().strip("|").split("|")]
    while cells and cells[0] == "":
        cells.pop(0)
    while cells and cells[-1] == "":
        cells.pop()
    return cells


def _drop_blank_columns(headers: list[str], rows: list[list[str]]) -> tuple[list[str], list[list[str]]]:
    """Remove columns that are empty in every header and data cell."""
    if not headers:
        return headers, rows
    keep = []
    for ci in range(len(headers)):
        col_vals = [headers[ci]] + [row[ci] if ci < len(row) else "" for row in rows]
        if any(v.strip() for v in col_vals):
            keep.append(ci)
    if len(keep) == len(headers):
        return headers, rows
    new_headers = [headers[i] for i in keep]
    new_rows = [[row[i] if i < len(row) else "" for i in keep] for row in rows]
    return new_headers, new_rows


def _table_column_widths(headers: list[str], usable_in: float) -> list[float]:
    """Narrow S.No. column; share remaining width across other columns."""
    if headers and headers[0] in ("S.No.", "#"):
        serial_w = min(0.45, usable_in * 0.06)
        rest_w = usable_in - serial_w
        if len(headers) == 6 and headers[1] == "Date":
            # S.No. | Date | Subject | To | Dispatch | File
            return [
                serial_w,
                rest_w * 0.09,
                rest_w * 0.34,
                rest_w * 0.14,
                rest_w * 0.18,
                rest_w * 0.25,
            ]
        if len(headers) == 9 and headers[1] == "Date" and headers[7] == "Attachments":
            return [
                serial_w,
                rest_w * 0.07,
                rest_w * 0.09,
                rest_w * 0.22,
                rest_w * 0.08,
                rest_w * 0.13,
                rest_w * 0.06,
                rest_w * 0.05,
                rest_w * 0.30,
            ]
        if len(headers) == 8 and headers[1] == "Date" and "Dispatch No" in headers[5]:
            return [
                serial_w,
                rest_w * 0.07,
                rest_w * 0.10,
                rest_w * 0.24,
                rest_w * 0.08,
                rest_w * 0.14,
                rest_w * 0.07,
                rest_w * 0.30,
            ]
        if len(headers) == 7 and headers[1] == "Date" and headers[4] == "Sent by":
            return [
                serial_w,
                rest_w * 0.08,
                rest_w * 0.11,
                rest_w * 0.30,
                rest_w * 0.09,
                rest_w * 0.14,
                rest_w * 0.28,
            ]
        if len(headers) == 3 and headers[-1] == "Location":
            return [serial_w, rest_w * 0.62, rest_w * 0.38]
        if len(headers) == 2:
            return [serial_w, rest_w]
        rest = rest_w / max(len(headers) - 1, 1)
        return [serial_w] + [rest] * (len(headers) - 1)
    return [usable_in / len(headers)] * len(headers)


def _set_column_widths(table, widths_inches: list[float]) -> None:
    """Assign fixed column widths so Word does not leave empty side space."""
    for row in table.rows:
        for ci, width in enumerate(widths_inches):
            if ci < len(row.cells):
                row.cells[ci].width = Inches(width)


def _add_formatted_runs(paragraph, text: str, *, italic_default: bool = False) -> None:
    """Add runs to a paragraph, handling **bold**, *italic*, and `code`."""
    pattern = re.compile(r"(\*\*[^*]+\*\*|\*[^*]+\*|`[^`]+`)")
    pos = 0
    for match in pattern.finditer(text):
        if match.start() > pos:
            run = paragraph.add_run(text[pos : match.start()])
            run.italic = italic_default
        token = match.group()
        if token.startswith("**"):
            run = paragraph.add_run(token[2:-2])
            run.bold = True
        elif token.startswith("`"):
            run = paragraph.add_run(token[1:-1])
            run.font.name = "Consolas"
            run._element.rPr.rFonts.set(qn("w:eastAsia"), "Consolas")
            run.font.size = Pt(9)
        else:
            run = paragraph.add_run(token[1:-1])
            run.italic = True
        pos = match.end()
    if pos < len(text):
        run = paragraph.add_run(text[pos:])
        run.italic = italic_default


def _set_table_borders(table) -> None:
    tbl = table._tbl
    tbl_pr = tbl.tblPr
    if tbl_pr is None:
        from docx.oxml import OxmlElement

        tbl_pr = OxmlElement("w:tblPr")
        tbl.insert(0, tbl_pr)
    borders = tbl_pr.find(qn("w:tblBorders"))
    if borders is None:
        from docx.oxml import OxmlElement

        borders = OxmlElement("w:tblBorders")
        tbl_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        element = borders.find(qn(f"w:{edge}"))
        if element is None:
            from docx.oxml import OxmlElement

            element = OxmlElement(f"w:{edge}")
            borders.append(element)
        element.set(qn("w:val"), "single")
        element.set(qn("w:sz"), "4")
        element.set(qn("w:color"), "AAAAAA")


def _set_section_landscape(section) -> None:
    """Switch a Word section to landscape (swap width and height)."""
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width, section.page_height = section.page_height, section.page_width


def markdown_to_docx(md_path: Path, docx_path: Path, *, landscape: bool = False) -> None:
    """Convert the generated markdown index into a formatted Word document."""
    lines = md_path.read_text(encoding="utf-8").splitlines()
    doc = Document()

    section = doc.sections[0]
    section.top_margin = Inches(0.8)
    section.bottom_margin = Inches(0.8)
    section.left_margin = Inches(0.9)
    section.right_margin = Inches(0.9)
    if landscape:
        _set_section_landscape(section)

    normal = doc.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(11)

    i = 0
    while i < len(lines):
        line = lines[i]

        if not line.strip():
            i += 1
            continue

        if line.strip() == "---":
            i += 1
            continue

        if line.startswith("# "):
            p = doc.add_heading(line[2:].strip(), level=0)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            i += 1
            continue

        if line.startswith("## "):
            doc.add_heading(line[3:].strip(), level=1)
            i += 1
            continue

        if line.startswith("### "):
            doc.add_heading(line[4:].strip(), level=2)
            i += 1
            continue

        if line.startswith("|") and i + 1 < len(lines) and lines[i + 1].startswith("|---"):
            headers = _parse_table_cells(line)
            i += 2
            rows: list[list[str]] = []
            while i < len(lines) and lines[i].startswith("|"):
                rows.append(_parse_table_cells(lines[i]))
                i += 1

            headers, rows = _drop_blank_columns(headers, rows)
            if not headers:
                continue

            table = doc.add_table(rows=1 + len(rows), cols=len(headers))
            table.style = "Table Grid"
            table.autofit = False
            table.allow_autofit = False
            _set_table_borders(table)

            usable_in = (
                section.page_width - section.left_margin - section.right_margin
            ) / 914400
            _set_column_widths(table, _table_column_widths(headers, usable_in))

            hdr_cells = table.rows[0].cells
            for col, header in enumerate(headers):
                hdr_cells[col].text = ""
                p = hdr_cells[col].paragraphs[0]
                run = p.add_run(header)
                run.bold = True
                shading = hdr_cells[col]._tc.get_or_add_tcPr()
                from docx.oxml import OxmlElement

                shd = OxmlElement("w:shd")
                shd.set(qn("w:fill"), "D9E2F3")
                shading.append(shd)

            for row_idx, row_data in enumerate(rows):
                row_cells = table.rows[row_idx + 1].cells
                for col, cell_text in enumerate(row_data):
                    row_cells[col].text = cell_text

            doc.add_paragraph()
            continue

        if line.startswith("**") and line.rstrip().endswith("**"):
            p = doc.add_paragraph()
            _add_formatted_runs(p, line.strip())
            i += 1
            continue

        if line.startswith("*") and line.endswith("*") and not line.startswith("**"):
            p = doc.add_paragraph()
            _add_formatted_runs(p, line.strip(), italic_default=True)
            i += 1
            continue

        p = doc.add_paragraph()
        _add_formatted_runs(p, line.strip())
        i += 1

    doc.save(docx_path)


def _section_short_title(title: str) -> str:
    short = title.split("—")[0].strip() if "—" in title else title
    if ". " in short[:4]:
        short = short.split(". ", 1)[-1]
    return short


def _escape_table_cell(text: str) -> str:
    return text.replace("|", "/").replace("\n", " ")


def _parse_ddmmyyyy_from_name(name: str) -> str | None:
    """Parse DDMMYYYY or DDMMYY from start of filename, or trailing DDMMYY (e.g. 241024)."""
    m = re.match(r"^(\d{2})(\d{2})(\d{4})", name)
    if m:
        return f"{m.group(1)}.{m.group(2)}.{m.group(3)}"
    m = re.match(r"^(\d{2})[.\s-](\d{2})[.\s-](\d{4})", name)
    if m:
        return f"{m.group(1)}.{m.group(2)}.{m.group(3)}"
    m = re.search(r"(\d{2})(\d{2})(\d{2})(?:\D*$)", name)
    if m:
        dd, mm, yy = m.group(1), m.group(2), m.group(3)
        year = f"20{yy}" if len(yy) == 2 else yy
        return f"{dd}.{mm}.{year}"
    return None


def _extract_date_from_md(md_path: Path) -> str | None:
    try:
        text = _read_md_text(md_path)
    except OSError:
        return None
    patterns = [
        r"\*\*Date:\*\*\s*(\d{1,2}\.\d{1,2}\.\d{4})",
        r"^Date:\s*(\d{1,2}\.\d{1,2}\.\d{4})",
        r"^Date:\s*(\d{1,2}\.\d{2}\.\d{4})",
    ]
    for pattern in patterns:
        match = re.search(pattern, text, re.MULTILINE | re.IGNORECASE)
        if match:
            return match.group(1)
    return None


def _date_sort_key(date_str: str) -> tuple[int, int, int, str]:
    """Sort key YYYY-MM-DD; unknown/draft dates sort last."""
    if not date_str or date_str in ("—", "-") or date_str.startswith("[") or date_str.upper().startswith("TBD"):
        return (9999, 12, 31, date_str)
    m = re.match(r"(\d{1,2})\.(\d{1,2})\.(\d{4})", date_str.strip())
    if m:
        d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        return (y, mo, d, date_str)
    return (9999, 12, 31, date_str)


def _date_heading(date_str: str) -> str:
    """Format date as section heading, e.g. 01 June 2026."""
    if not date_str or date_str.startswith("[") or date_str.upper().startswith("TBD"):
        return date_str
    m = re.match(r"(\d{1,2})\.(\d{1,2})\.(\d{4})", date_str.strip())
    if not m:
        return date_str
    d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
    try:
        return datetime(y, mo, d).strftime("%d %B %Y")
    except ValueError:
        return date_str


def _extract_dispatch_ref(name: str) -> str | None:
    m = re.search(r"Document\s*(\d+)", name, re.I)
    if m:
        return f"DTCP Diary Dispatch No. {m.group(1)}"
    if "speedpost" in name.lower():
        d = re.search(r"(\d{1,2})-(\d{1,2})-(\d{2,4})", name)
        if d:
            yr = d.group(3)
            if len(yr) == 2:
                yr = "20" + yr
            return f"Speed Post receipt — {d.group(1)}.{d.group(2)}.{yr}"
        return "Speed Post receipt (see filename)"
    return None


def _extract_tracking_from_name(name: str) -> str | None:
    """India Post article number, e.g. ED212696853IN."""
    m = re.search(r"\b([A-Z]{2}\d{9}[A-Z]{2})\b", name, re.I)
    return m.group(1).upper() if m else None


def _date_to_ddmmyyyy_key(date_str: str) -> str | None:
    m = re.match(r"(\d{2})\.(\d{2})\.(\d{4})", date_str.strip())
    return f"{m.group(1)}{m.group(2)}{m.group(3)}" if m else None


def _valid_date_keys_in_name(name: str) -> list[str]:
    """DDMMYYYY keys in a filename (not digits from tracking numbers)."""
    keys: list[str] = []
    for match in re.finditer(r"(?<!\d)(\d{2})(\d{2})(\d{4})(?!\d)", name):
        day, month, year = int(match.group(1)), int(match.group(2)), int(match.group(3))
        if 1 <= day <= 31 and 1 <= month <= 12 and 2010 <= year <= 2035:
            keys.append(match.group(0))
    return keys


def _build_tracking_index() -> dict[str, list[str]]:
    """Map letter date key (DDMMYYYY) to tracking numbers from dated proof files."""
    index: dict[str, list[str]] = defaultdict(list)
    for path in _speedpost_proof_files():
        tracking = _extract_tracking_from_file(path)
        if not tracking:
            continue
        for date_key in _valid_date_keys_in_name(path.name):
            if tracking not in index[date_key]:
                index[date_key].append(tracking)
    return index


def _extract_tracking_from_file(path: Path) -> str:
    """Read India Post article number from filename or file body."""
    found = _extract_tracking_from_name(path.name)
    if found:
        return found
    if path.suffix.lower() != ".pdf":
        return ""
    try:
        raw = path.read_bytes().decode("latin-1", errors="ignore")
        compact = re.sub(r"\s+", "", raw)
        match = re.search(r"([A-Z]{2}\d{9}[A-Z]{2})", compact, re.I)
        if match:
            return match.group(1).upper()
    except OSError:
        pass
    return ""


def _speedpost_proof_files() -> list[Path]:
    found: list[Path] = []
    for path in root.rglob("*"):
        if not path.is_file() or should_skip(path):
            continue
        name_lower = path.name.lower()
        if not any(
            x in name_lower
            for x in ("speedpost", "speed post", "consignment", "tracking")
        ):
            continue
        if path.suffix.lower() in {".pdf", ".jpg", ".jpeg", ".png"}:
            found.append(path)
    return found


def _proof_matches_letter(path: Path, date: str, stem: str) -> bool:
    letter_key = _date_to_ddmmyyyy_key(date)
    path_dates = _valid_date_keys_in_name(path.name)
    if letter_key and letter_key in path_dates:
        return True
    stem_key = re.sub(r"\D", "", stem)[:8]
    if stem_key and len(stem_key) >= 8 and stem_key in re.sub(r"\D", "", path.name):
        return True
    path_lower = path.name.lower()
    if "speedpost" in path_lower or "speed post" in path_lower:
        if "execution" in stem.lower() and "execution" in path_lower:
            if letter_key == "24102024" or "241024" in re.sub(r"\D", "", stem):
                return True
    return False


def _consignment_tracking_for_proof(
    date_key: str, matched_proofs: list[Path]
) -> str:
    """Link consignment PDF to a letter when a dated speed-post proof is on file."""
    if not date_key or not matched_proofs:
        return ""
    if not any(date_key in re.sub(r"\D", "", p.name) for p in matched_proofs):
        return ""
    for path in sorted(root.rglob("Consignment*")):
        if path.is_file():
            tracking = _extract_tracking_from_name(path.name)
            if tracking:
                return tracking
    return ""


def _folder_speedpost_fields(
    date: str,
    stem: str,
    entry: dict[str, str],
    dispatch: str,
) -> tuple[str, str]:
    """Speed-post tracking and done — folder proof only; blank if none."""
    if not _letter_used_speedpost(dispatch, entry):
        return "", ""
    date_key = _date_to_ddmmyyyy_key(date) or ""
    matched = [
        p for p in _speedpost_proof_files() if _proof_matches_letter(p, date, stem)
    ]
    if not matched:
        return "", ""
    for proof in matched:
        tracking = _extract_tracking_from_file(proof)
        if tracking:
            return tracking, "Yes"
    consignment = _consignment_tracking_for_proof(date_key, matched)
    if consignment:
        return consignment, "Yes"
    return "", ""


def _letter_used_speedpost(dispatch: str, entry: dict[str, str]) -> bool:
    d = dispatch.lower()
    return (
        "speed post" in d
        or "speedpost" in d
        or entry.get("speedpost_required") == "yes"
    )


def _speedpost_receipt_pending(
    entry: dict[str, str],
    dispatch: str,
    tracking_no: str,
    date: str,
    stem: str,
) -> bool:
    """True when speed post expected but no proof file is linked in the folder."""
    if not _letter_used_speedpost(dispatch, entry):
        return False
    if tracking_no:
        return False
    matched = [
        p for p in _speedpost_proof_files() if _proof_matches_letter(p, date, stem)
    ]
    return len(matched) == 0


DISPATCH_TABLE_HEADER = (
    "| S.No. | Date | To | Subject | Sent by | "
    "Speed Post Tracking / Dispatch No. | Speed Post Done? | Attachments | File |"
)
DISPATCH_TABLE_RULE = (
    "|-------|------|-----|---------|---------|"
    "-----------------------------------|------------------|-------------|------|"
)


def _find_companion_md(stem: str, *, for_dispatch: bool = False) -> Path | None:
    candidates: list[Path] = []
    skip_parts = set() if for_dispatch else SKIP_PATH_PARTS
    for path in root.rglob(f"{stem}.md"):
        if should_skip(path) or any(p in skip_parts for p in path.parts):
            continue
        candidates.append(path)
    if not candidates:
        return None
    return sorted(candidates, key=lambda p: (len(p.parts), str(p).lower()))[0]


def _clean_subject(text: str) -> str:
    text = re.sub(r"\*\*", "", text)
    text = re.sub(r"\s+", " ", text).strip()
    if len(text) > 220:
        text = text[:217].rstrip() + "..."
    return text


def _read_md_text(path: Path) -> str:
    raw = path.read_bytes()
    for encoding in ("utf-8-sig", "utf-8", "utf-16", "utf-16-le", "latin-1"):
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


def _enclosures_section_text(text: str) -> str | None:
    """Return body text after an enclosures / annexures heading."""
    patterns = [
        r"^##[^\n]*(?:Enclos(?:ure)?s?|Annexures?)[^\n]*\n",
        r"^###\s+Enclosures?\s*\n",
        r"^\*\*\d+\.\s+Enclosures?\*\*[^\n]*\n",
        r"^Annexures\s*\n",
    ]
    best_start: int | None = None
    best_end = 0
    for pattern in patterns:
        match = re.search(pattern, text, re.MULTILINE | re.IGNORECASE)
        if match and (best_start is None or match.start() < best_start):
            best_start = match.start()
            best_end = match.end()
    if best_start is None:
        return None
    section = text[best_end:]
    stop = re.search(
        r"\n---\n|\n\*\*Place:\*\*|\n\*\*Filing note|\n\*\*Applicant\*\*|\n\*\*From,\*\*|\n##\s+(?!#)",
        section,
        re.IGNORECASE,
    )
    if stop:
        section = section[: stop.start()]
    return section


def _count_items_in_enclosures_section(section: str) -> int:
    count = 0
    for line in section.splitlines():
        stripped = line.strip()
        if not stripped:
            continue
        if stripped.startswith("|") and not re.match(r"^\|[\s\-:|]+\|$", stripped):
            cells = [_strip_md_inline(c) for c in stripped.strip("|").split("|")]
            if cells and not re.match(
                r"^(s\.?\s*no\.?|#|annexure|encl\.?|--|-)$", cells[0], re.I
            ):
                count += 1
            continue
        if re.match(
            r"^\*\*(?:Encl\.?\s*[\w.]+|\d+(?:\.\d+(?:\([a-z]\))?)?\.?)\*\*",
            stripped,
            re.I,
        ):
            count += 1
            continue
        if re.match(r"^Annexure\s+\d+\s*[-–]", stripped, re.I):
            count += 1
            continue
        if re.match(r"^\d+\.\s+", stripped) and re.search(
            r"annexure|encl\.|`", stripped, re.I
        ):
            count += 1
            continue
        if re.match(r"^\d+\.\s+", stripped):
            count += 1
    return count


def _extract_attachment_count_from_md(md_path: Path) -> str:
    try:
        text = _read_md_text(md_path)
    except OSError:
        return ""

    head = text[:2500]
    for pattern in (
        r"\*\*Attachments\s*\((\d+)\)\*\*",
        r"Attachments\s*\((\d+)\)",
        r"\*\*(\d+)\s+attachments?\*\*",
        r"\*\*Note:\*\*[^\n]*?(\d+)\s+attachments?\b",
    ):
        match = re.search(pattern, head, re.I)
        if match:
            return match.group(1)

    section = _enclosures_section_text(text)
    if section:
        count = _count_items_in_enclosures_section(section)
        if count:
            return str(count)

    block = re.search(
        r"\*\*\d+\.\s+Enclosures?\*\*[^\n]*\n(.*?)(?=\n\*\*|\n---|\n##|\Z)",
        text,
        re.MULTILINE | re.DOTALL | re.I,
    )
    if block:
        count = _count_items_in_enclosures_section(block.group(1))
        if count:
            return str(count)

    encl_lines = re.findall(
        r"^\*\*Encl\.?\s*\d+\*\*",
        text,
        re.MULTILINE | re.I,
    )
    if encl_lines:
        return str(len(encl_lines))

    return ""


def _resolve_attachment_count(entry: dict[str, str], md: Path | None) -> str:
    if entry.get("attachments") is not None:
        return str(entry["attachments"])
    if md:
        return _extract_attachment_count_from_md(md)
    return ""


def _extract_subject_from_md(md_path: Path) -> str | None:
    try:
        text = _read_md_text(md_path)
    except OSError:
        return None
    patterns = [
        r"## Subject\s*\n+\s*\*\*(.+?)\*\*",
        r"## Subject\s*\n+\s*(.+?)(?:\n\n|\n---)",
        r"\*\*Subject:\*\*\s*(.+?)(?:\n\n|\n---|\n\n---)",
        r"^Subject:\s*(.+?)(?:\n\n|\n---)",
    ]
    for pattern in patterns:
        match = re.search(pattern, text, re.MULTILINE | re.DOTALL | re.IGNORECASE)
        if match:
            subject = match.group(1).strip()
            subject = subject.split("\n")[0].strip()
            return _clean_subject(subject)
    return None


def _resolve_dispatch_file(stem: str, explicit: str | None = None) -> str:
    if explicit and (root / explicit).exists():
        return explicit.replace("\\", "/")
    matches = [
        p
        for p in root.rglob(f"{stem}*")
        if p.is_file()
        and not should_skip(p)
        and p.suffix.lower() in {".pdf", ".docx", ".doc"}
        and " (1)" not in p.name
    ]
    if not matches:
        return stem
    best = sorted(matches, key=lambda p: (len(p.parts), str(p).lower()))[0]
    return str(best.relative_to(root)).replace("\\", "/")


# Outgoing correspondence register — chronological (allottee submissions).
# Subjects enriched from companion .md where available; dispatch nos. from receipts/filenames.
DISPATCH_REGISTRY: list[dict[str, str]] = [
    {
        "sort": "20241024",
        "date": "24.10.2024",
        "to": "HARERA, Gurugram",
        "subject": "Execution complaint / application — Complaint No. 4463 of 2021 — Unit 3051",
        "dispatch": "Speed Post — posted 11.10.2024",
        "file_stem": "3_ATS_Execution241024 (1)",
        "file": (
            "Change_of_devloper_by_ATS/Reply by ATS/reply by Allotee/"
            "DTPC_HRERA_CONDITIONAL_APPROVAL/3_ATS_Execution241024 (1).pdf"
        ),
    },
    {
        "sort": "20250530",
        "date": "30.05.2026",
        "to": "ATS / Hasta / Homekraft",
        "subject": "Objection — suspected back-dated intimation letter dated 07.05.2026; demand proof of service",
        "dispatch": "Email / legal notice",
        "file_stem": "ATS_Objection_Intimation_Letter",
        "file": "Change_of_devloper_by_ATS/Reply by ATS/reply by Allotee/Fabrication/ATS_Objection_Intimation_Letter.md",
    },
    {
        "sort": "20250601",
        "date": "01.06.2026",
        "to": "HARERA, Gurugram",
        "subject": "Application for additional submission / rejoinder in pending execution proceedings (response to ATS reply May 2026)",
        "dispatch": "HARERA execution portal / email",
        "file_stem": "01062026_Application_Additional_Submission_in_Pending_Execution_Unit_3051",
    },
    {
        "sort": "20250601b",
        "date": "01.06.2026",
        "to": "HARERA, Gurugram",
        "subject": "Rejoinder to ATS response — protection of BBA rights and HARERA order dated 18.04.2024",
        "dispatch": "HARERA execution portal / email",
        "file_stem": "01062026_Rejoinder_ATS_REPLY",
    },
    {
        "sort": "20250601c",
        "date": "01.06.2026",
        "to": "HARERA + DTCP",
        "subject": "Request that COD / revised-plan approval be conditional on executing HARERA order 18.04.2024",
        "dispatch": "Email — rera@hareraggm.gov.in; tcpharyana7@gmail.com; stp4.gurugram.tcp@gmail.com",
        "file_stem": "01062026_Covering_Letter_Conditional_Approval",
        "file": "Change_of_devloper_by_ATS/Reply by ATS/reply by Allotee/DTPC_HRERA_CONDITIONAL_APPROVAL/md_files/01062026_Covering_Letter_Conditional_Approval.md",
    },
    {
        "sort": "20250606",
        "date": "06.06.2026",
        "to": "HARERA + DTCP",
        "subject": "Conditional approval request — execute HARERA order before COD / revised plans",
        "dispatch": "Email — HARERA & DTCP (see Gmail export in folder)",
        "file_stem": "06062026_Covering_Letter_Conditional_Approval",
    },
    {
        "sort": "20250613",
        "date": "13.06.2026",
        "to": "HARERA + DTCP",
        "subject": "Objection to ATS joint affidavit 08.06.2026; buyer safeguards before COD approval",
        "dispatch": "Email — government authorities",
        "file_stem": "13062026_Covering_Note_Government_Authorities_Conditional_Approval",
    },
    {
        "sort": "20250615",
        "date": "15.06.2026",
        "to": "HARERA, Gurugram",
        "subject": "Composite execution application — DTCP revised plan, COD, affidavit gaps, impleadment",
        "dispatch": "HARERA execution proceedings",
        "file_stem": "15062026_Composite_Execution_Application_Gap_Analysis_Covering_Note_Unit_3051",
    },
    {
        "sort": "20250618",
        "date": "18.06.2026",
        "to": "DTCP STP-4, Gurugram",
        "subject": "Filing / service proof — conditional approval submissions",
        "dispatch": "DTCP Diary Dispatch No. 86",
        "file_stem": "18062026_diary_dispatch_Document 86_pdf",
    },
    {
        "sort": "20250619",
        "date": "19.06.2026",
        "to": "DTCP STP-4, Gurugram",
        "subject": "Post-hearing representation on MOM dated 18.06.2026 — revised building plan ZP-938-II",
        "dispatch": "Email — stp4.gurugram.tcp@gmail.com",
        "file_stem": "19062026_Rejoinder_MOM_Hearing_18_06_2026_Unit_3051",
    },
    {
        "sort": "20250622",
        "date": "22.06.2026",
        "to": "DTCP STP-4, Gurugram",
        "subject": "Further filing / service proof — post-hearing papers",
        "dispatch": "DTCP Diary Dispatch No. 92",
        "file_stem": "22062026 Diary dispatch number Document 92",
    },
    {
        "sort": "20250623a",
        "date": "23.06.2026",
        "to": "PNBHFL, Gurugram",
        "subject": "Procedure & document checklist — direct pre-EMI payment by builder into loan account",
        "dispatch": "Email + Speed Post",
        "speedpost_required": "yes",
        "file_stem": "23062026_PNB_PreEMI_Procedure_Unit_3051",
    },
    {
        "sort": "20250623b",
        "date": "23.06.2026",
        "to": "ATS / Homekraft",
        "subject": "Reply re PNBHFL IFSC / pre-EMI payment — Unit 3051",
        "dispatch": "Email — grandstand.customercare@homekraft.in",
        "file_stem": "23062026_ATS_IFSC_PreEMI_Unit_3051_FINAL",
    },
    {
        "sort": "20250623c",
        "date": "23.06.2026",
        "to": "ATS / Homekraft",
        "subject": "Reply to ATS email dated 22.06.2026 — pre-EMI meeting / HARERA compliance",
        "dispatch": "Email",
        "file_stem": "23062026_Reply_to_ATS_PreEMI_Meeting_Unit_3051",
    },
    {
        "sort": "20250623d",
        "date": "23.06.2026",
        "to": "HARERA, Gurugram",
        "subject": "Execution application — builder ECS/NACH, allottee ECS stand-down, CIBIL protection",
        "dispatch": "HARERA execution proceedings",
        "file_stem": "23062026_RERA_Execution_Application_Builder_ECS_Allottee_CIBIL_Unit_3051",
    },
    {
        "sort": "20250624a",
        "date": "24.06.2026",
        "to": "PNBHFL, Gurugram",
        "subject": "Authorisation letter — Mr Lalit — branch visit for ECS / pre-EMI compliance",
        "dispatch": "Email — gurgaon@pnbhousing.com",
        "file_stem": "24062026_Authorisation_Letter_Mr_Lalit_PNB_Branch_Visit_Unit_3051",
    },
    {
        "sort": "20250625a",
        "date": "25.06.2026",
        "to": "ATS / Homekraft",
        "subject": "Reply to colonizer tabulated response dated 24.06.2026 — Unit 3051 objections",
        "dispatch": "Email",
        "file_stem": "25062026_Reply_to_ATS_24062026_Objections_Unit_3051",
    },
    {
        "sort": "20250625b",
        "date": "25.06.2026",
        "to": "DTCP + HARERA",
        "subject": "Submission for record — affidavit-cum-undertaking and consent letter dated 25.06.2026",
        "dispatch": "DTCP / HARERA file",
        "file_stem": "25062026_Affidavit_cum_undertaking",
        "attachments": "2",
    },
    {
        "sort": "20250625c",
        "date": "25.06.2026",
        "to": "PNBHFL, Gurugram",
        "subject": "Request for direct pre-EMI payment by builder/promoter into loan account",
        "dispatch": "Email (Gmail export on file)",
        "file_stem": "Gmail_23062026_REQUEST_FOR_DIRECT_EMI_PAYMENT_BY_BUILDER",
        "md_stem": "25062026_Gmail_REQUEST_Direct_PreEMI_Payment_Builder_Unit_3051",
    },
    {
        "sort": "20250629",
        "date": "29.06.2026",
        "to": "DTCP + HARERA",
        "subject": "Clarification — consent affidavit & undertaking dated 25.06.2026 to be read in conjunction",
        "dispatch": "DTCP STP-4 / HARERA email",
        "file_stem": "29062026_Clarification_Consent_Affidavit_Unit_3051",
    },
]


def _build_dispatch_rows() -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []
    for entry in DISPATCH_REGISTRY:
        stem = entry["file_stem"]
        md_stem = entry.get("md_stem", stem)
        md = _find_companion_md(md_stem, for_dispatch=True)
        subject = entry.get("subject", "")
        if md:
            extracted = _extract_subject_from_md(md)
            if extracted and len(extracted) <= 180 and "Dear Sir" not in extracted:
                subject = extracted

        file_path = entry.get("file") or _resolve_dispatch_file(stem)
        if file_path.endswith(".md"):
            docx_alt = _resolve_dispatch_file(stem)
            if not docx_alt.endswith(".md") and docx_alt != stem:
                file_path = docx_alt
            elif entry.get("file"):
                file_path = entry["file"]

        dispatch = entry.get("dispatch", "")
        if not dispatch or dispatch.startswith("["):
            from_name = Path(file_path).name
            auto = _extract_dispatch_ref(from_name)
            if auto:
                dispatch = auto

        date = entry.get("date", "")
        if md:
            md_date = _extract_date_from_md(md)
            if md_date:
                date = md_date
        if not date or date.startswith("[") or date.upper().startswith("TBD"):
            parsed = (
                _parse_ddmmyyyy_from_name(stem)
                or _parse_ddmmyyyy_from_name(Path(file_path).name)
            )
            if parsed:
                date = parsed

        tracking_no, speedpost_done = _folder_speedpost_fields(
            date, stem, entry, dispatch
        )

        rows.append(
            {
                "sort": entry["sort"],
                "date": date,
                "date_key": _date_sort_key(date),
                "status": entry.get("status", "SENT"),
                "subject": subject,
                "to": entry["to"],
                "dispatch": dispatch,
                "tracking_no": tracking_no,
                "speedpost_done": speedpost_done,
                "speedpost_receipt_pending": _speedpost_receipt_pending(
                    entry, dispatch, tracking_no, date, stem
                ),
                "attachments": _resolve_attachment_count(entry, md),
                "file": file_path,
            }
        )
    sent_rows = [r for r in rows if r.get("status") != "DRAFT"]
    draft_rows = [r for r in rows if r.get("status") == "DRAFT"]
    return sorted(sent_rows, key=lambda r: (r["date_key"], r["sort"]), reverse=True) + draft_rows


def _short_subject(text: str, limit: int = 72) -> str:
    text = re.sub(r"\s+", " ", text).strip()
    if len(text) <= limit:
        return text
    return text[: limit - 3].rstrip() + "..."


def _sent_by_label(row: dict[str, str]) -> str:
    dispatch = row.get("dispatch", "")
    d = dispatch.lower()
    if "diary dispatch" in d:
        return "DTCP Diary"
    if "harera file" in d or "dtcp / harera file" in d:
        return "Filed"
    if "harera execution" in d or "harera portal" in d:
        return "HARERA"
    if "legal notice" in d:
        return "Email / Notice"
    if "speed post" in d or "speedpost" in d:
        return "Speed Post"
    if "email" in d:
        return "Email"
    if "stp-4" in d or "dtcp" in d:
        return "Email"
    return ""


def _tracking_dispatch_ref(row: dict[str, str]) -> str:
    """Speed post tracking or DTCP diary dispatch number from folder / dispatch."""
    tracking = row.get("tracking_no", "")
    if tracking:
        return tracking
    dispatch = row.get("dispatch", "")
    match = re.search(r"Diary Dispatch\s*No\.\s*(\d+)", dispatch, re.I)
    if match:
        return f"Diary {match.group(1)}"
    return ""


def _dispatch_table_row(sn: int, row: dict[str, str]) -> str:
    return (
        "| {sn} | {date} | {to} | {subject} | {sent_by} | "
        "{ref} | {done} | {attachments} | `{file}` |"
    ).format(
        sn=sn,
        date=_escape_table_cell(row["date"]),
        to=_escape_table_cell(row["to"]),
        subject=_escape_table_cell(_short_subject(row["subject"])),
        sent_by=_escape_table_cell(_sent_by_label(row)),
        ref=_escape_table_cell(_tracking_dispatch_ref(row)),
        done=_escape_table_cell(row.get("speedpost_done", "")),
        attachments=_escape_table_cell(row.get("attachments", "")),
        file=Path(row["file"]).name,
    )


def generate_dispatch_correspondence_index() -> None:
    """Build a short outgoing correspondence register."""
    all_rows = _build_dispatch_rows()
    rows = [r for r in all_rows if r.get("status") != "DRAFT"]
    drafts = [r for r in all_rows if r.get("status") == "DRAFT"]
    today = datetime.now().strftime("%d %B %Y")
    latest = rows[0]["date"] if rows else "—"
    speedpost_pending = [r for r in rows if r.get("speedpost_receipt_pending")]

    lines: list[str] = [
        "# Dispatch & Correspondence Index",
        "",
        f"**Unit 3051** · HARERA Complaint **4463/2021** · **{len(rows)} letters** · "
        f"Latest: **{latest}** · Updated: {today}",
        "",
        "---",
        "",
        "## Outgoing Register (newest first)",
        "",
        DISPATCH_TABLE_HEADER,
        DISPATCH_TABLE_RULE,
    ]

    for i, row in enumerate(rows, 1):
        lines.append(_dispatch_table_row(i, row))

    if speedpost_pending:
        lines += ["", "---", "", "## Pending", ""]
        for row in speedpost_pending:
            lines.append(
                f"- **{row['date']}** — add speed-post proof to folder — "
                f"`{Path(row['file']).name}`"
            )

    if drafts:
        lines += [
            "",
            "---",
            "",
            "## Drafts (not sent)",
            "",
            DISPATCH_TABLE_HEADER,
            DISPATCH_TABLE_RULE,
        ]
        for i, row in enumerate(drafts, len(rows) + 1):
            lines.append(_dispatch_table_row(i, row))

    lines += [
        "",
        "---",
        "",
        "*Regenerate: `python generate_index.py`*",
    ]

    dispatch_md.write_text("\n".join(lines), encoding="utf-8")
    markdown_to_docx(dispatch_md, dispatch_docx, landscape=True)
    print(f"Written {dispatch_md.name} — {len(rows)} correspondence entries")
    print(f"Written {dispatch_docx.name}")


def generate_main_files_index() -> None:
    """Build deduplicated index of main documents only (no markdown)."""
    files_by_dir, raw_count, duplicate_copies = collect_unique_main_files()
    unique_count = sum(len(files) for files in files_by_dir.values())
    today = datetime.now().strftime("%d %B %Y")

    lines: list[str] = [
        "# Main Files Index",
        "",
        "**Project:** ATS Grandstand Phase-I, Sector 99A, Gurugram  ",
        "**Unit:** 3051, Tower 03  ",
        "**Complaint:** HARERA Complaint No. 4463 of 2021  ",
        f"**Index generated:** {today}  ",
        f"**Unique main files:** {unique_count}  ",
        f"**Duplicate copies excluded:** {duplicate_copies} (from {raw_count} main-file copies scanned)",
        "",
        "*Includes PDF, Word, Excel, and image files only. Excludes markdown, scripts, Google shortcuts, and duplicate copies in annexure folders.*",
        "",
        "---",
        "",
        "## Quick Navigation",
        "",
        "| S.No. | Section | Folder | Files |",
        "|-------|---------|--------|-------|",
    ]

    sec_num = 0
    for path, title, _ in MAIN_SECTIONS:
        cnt = len(files_by_dir.get(path, []))
        if cnt == 0:
            continue
        sec_num += 1
        folder = f"`{path}`" if path != "(ROOT)" else "*(root)*"
        lines.append(f"| {sec_num} | {_section_short_title(title)} | {folder} | {cnt} |")

    lines += ["", "---", ""]

    for path, title, desc in MAIN_SECTIONS:
        files = files_by_dir.get(path, [])
        if not files:
            continue

        lines.append(f"## {title}")
        if desc:
            lines.append(f"*{desc}*")
        lines.append("")

        location_label = "*(root)*" if path == "(ROOT)" else f"`{path}`"

        if path == "(ROOT)":
            cats: dict[str, list[str]] = defaultdict(list)
            for name in files:
                cats[categorize_root(name)].append(name)
            for cat in ROOT_CAT_ORDER:
                if not cats[cat]:
                    continue
                lines.append(f"### {cat}")
                lines.append("")
                lines.append("| S.No. | File | Location |")
                lines.append("|-------|------|----------|")
                for i, name in enumerate(sorted(cats[cat], key=str.lower), 1):
                    lines.append(f"| {i} | `{name}` | {location_label} |")
                lines.append("")
        else:
            lines.append("| S.No. | File | Location |")
            lines.append("|-------|------|----------|")
            for i, name in enumerate(files, 1):
                lines.append(f"| {i} | `{name}` | {location_label} |")
            lines.append("")

    covered = {p for p, _, _ in MAIN_SECTIONS}
    extra = sorted(set(files_by_dir.keys()) - covered)
    if extra:
        lines.append("## Other Folders")
        lines.append("")
        for epath in extra:
            files = files_by_dir[epath]
            lines.append(f"### `{epath}`")
            lines.append("")
            lines.append("| S.No. | File | Location |")
            lines.append("|-------|------|----------|")
            for i, name in enumerate(files, 1):
                lines.append(f"| {i} | `{name}` | `{epath}` |")
            lines.append("")

    lines += [
        "---",
        "",
        "## File Type Legend",
        "",
        "| S.No. | Extension | Type |",
        "|-------|-----------|------|",
        "| 1 | `.pdf` | Scanned documents, orders, letters |",
        "| 2 | `.docx` | Word documents for filing / sending |",
        "| 3 | `.xlsx` | Payment / dues spreadsheets |",
        "| 4 | `.jpeg` / `.jpg` | Payment cheque photos |",
        "| 5 | `.png` | Images / screenshots |",
        "",
        "---",
        "",
        "*To regenerate all indexes, run: `python generate_index.py`*",
    ]

    main_md.write_text("\n".join(lines), encoding="utf-8")
    markdown_to_docx(main_md, main_docx)
    print(f"Written {main_md.name} — {unique_count} unique main files ({duplicate_copies} duplicates excluded)")
    print(f"Written {main_docx.name}")
    generate_dispatch_correspondence_index()


def main() -> None:
    files_by_dir: dict[str, list[str]] = defaultdict(list)
    total = 0

    for f in sorted(root.rglob("*")):
        if not f.is_file() or should_skip(f):
            continue
        if f.name.startswith("00_MASTER_FILE_INDEX") or f.name in (
            "generate_index.py",
            "desktop.ini",
            "Ats execution pdf 241024.pdf",
        ):
            continue
        rel = f.relative_to(root)
        parent = str(rel.parent).replace("\\", "/") if str(rel.parent) != "." else "(ROOT)"
        files_by_dir[parent].append(f.name)
        total += 1

    lines: list[str] = []
    today = datetime.now().strftime("%d %B %Y")

    lines += [
        "# Master File Index",
        "",
        "**Project:** ATS Grandstand Phase-I, Sector 99A, Gurugram  ",
        "**Unit:** 3051, Tower 03  ",
        "**Complaint:** HARERA Complaint No. 4463 of 2021  ",
        f"**Index generated:** {today}  ",
        f"**Total files indexed:** {total}",
        "",
        "---",
        "",
        "## Quick Navigation",
        "",
        "| S.No. | Section | Folder | Files |",
        "|-------|---------|--------|-------|",
    ]

    sec_num = 0
    for path, title, _ in SECTIONS:
        cnt = len(files_by_dir.get(path, []))
        if cnt == 0:
            continue
        sec_num += 1
        folder = f"`{path}`" if path != "(ROOT)" else "*(root)*"
        lines.append(f"| {sec_num} | {_section_short_title(title)} | {folder} | {cnt} |")

    lines += [
        "",
        "### Detailed sub-indexes",
        "",
        "| S.No. | Document | Path |",
        "|-------|----------|------|",
        "| 1 | PNB attachments 101–114 | `PNB_ATS/README.md` |",
        "| 2 | PNB attachment descriptions | `PNB_ATS/attachments/README.md` |",
        "| 3 | Filing annexures (01.06.2026) | `Change_of_devloper_by_ATS/Reply by ATS/reply by Allotee/02062026_Annexures/00_ANNEXURE_INDEX.md` |",
        "",
        "---",
        "",
    ]

    for path, title, desc in SECTIONS:
        files = files_by_dir.get(path, [])
        if not files:
            continue

        lines.append(f"## {title}")
        if desc:
            lines.append(f"*{desc}*")
        lines.append("")

        if path == "(ROOT)":
            cats: dict[str, list[str]] = defaultdict(list)
            for name in sorted(files):
                cats[categorize_root(name)].append(name)
            for cat in ROOT_CAT_ORDER:
                if not cats[cat]:
                    continue
                lines.append(f"### {cat}")
                lines.append("")
                lines.append("| S.No. | File |")
                lines.append("|-------|------|")
                for i, name in enumerate(cats[cat], 1):
                    lines.append(f"| {i} | `{name}` |")
                lines.append("")
        else:
            lines.append("| S.No. | File |")
            lines.append("|-------|------|")
            for i, name in enumerate(sorted(files), 1):
                lines.append(f"| {i} | `{name}` |")
            lines.append("")

    covered = {p for p, _, _ in SECTIONS}
    extra = sorted(set(files_by_dir.keys()) - covered)
    if extra:
        lines.append("## Other Folders")
        lines.append("")
        for epath in extra:
            files = files_by_dir[epath]
            lines.append(f"### `{epath}`")
            lines.append("")
            lines.append("| S.No. | File |")
            lines.append("|-------|------|")
            for i, name in enumerate(sorted(files), 1):
                lines.append(f"| {i} | `{name}` |")
            lines.append("")

    lines += [
        "---",
        "",
        "## File Type Legend",
        "",
        "| S.No. | Extension | Type |",
        "|-------|-----------|------|",
        "| 1 | `.pdf` | Scanned documents, orders, letters |",
        "| 2 | `.docx` | Word drafts for filing / sending |",
        "| 3 | `.md` | Markdown text (editable source) |",
        "| 4 | `.xlsx` / `.gsheet` | Payment / dues spreadsheets |",
        "| 5 | `.jpeg` / `.jpg` | Payment cheque photos |",
        "| 6 | `.gdoc` | Google Docs shortcuts |",
        "| 7 | `.py` | Conversion / organization scripts |",
        "",
        "*Excludes `.git` internals and `__pycache__`.*",
        "",
        "---",
        "",
        "*To regenerate this index, run: `python generate_index.py`*",
    ]

    out_md.write_text("\n".join(lines), encoding="utf-8")
    markdown_to_docx(out_md, out_docx)
    print(f"Written {out_md.name} — {total} files indexed")
    print(f"Written {out_docx.name}")
    generate_main_files_index()


if __name__ == "__main__":
    main()
