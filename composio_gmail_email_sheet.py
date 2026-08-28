"""Build email date/time sheet from your LIVE Gmail inbox (last N months).

Reads directly from prof.vijyant.agarwal@gmail.com — not from local PDF exports.

Two ways to connect:
  A) Composio (default): COMPOSIO_API_KEY in composio.env + one-time --connect OAuth
  B) Gmail IMAP (direct): GMAIL_APP_PASSWORD in gmail.env (Google App Password)

Outputs:
  03_EMAIL_DATE_TIME_SHEET.md / .xlsx / .docx
"""
from __future__ import annotations

import argparse
import email
import imaplib
import json
import os
import re
import sys
import time
import webbrowser
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from email.header import decode_header
from email.utils import parseaddr, parsedate_to_datetime
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent
IST = timezone(timedelta(hours=5, minutes=30))
MY_EMAIL = "prof.vijyant.agarwal@gmail.com"
USER_ID = MY_EMAIL
LOCAL_ENV = ROOT / "composio.env"
GMAIL_ENV = ROOT / "gmail.env"

OUT_MD = ROOT / "03_EMAIL_DATE_TIME_SHEET.md"
OUT_XLSX = ROOT / "03_EMAIL_DATE_TIME_SHEET.xlsx"
OUT_DOCX = ROOT / "03_EMAIL_DATE_TIME_SHEET.docx"

EXACT_EMAILS = {
    "grandstand.customercare@homekraft.in",
    "rajeev.gupta@homekraft.in",
    "reajeev.gupta@homekraft.in",
    "stp4.gurugram.tcp@gmail.com",
    "tcpharyana7@gmail.com",
    "rera@hareraggm.gov.in",
    "udaivir.anand@atsgreen.com",
    "email@atsgreens.com",
    "email@atsgreeens.com",
    "customercare@pnbhousing.com",
    "gurgaon@pnbhousing.com",
    "krishan.kumar@pnbhousing.com",
}

DOMAIN_RULES: list[tuple[str, tuple[str, ...]]] = [
    ("ATS / Homekraft", ("homekraft.in", "atsgreen.com", "atsgreens.com", "atsgreeens.com")),
    ("DTCP", ("tcpharyana", "dtcp", "townandcountryplanning", "stp4.gurugram")),
    ("PNB Housing", ("pnbhousing.com",)),
    ("HARERA", ("hareraggm.gov.in", "hrera")),
]

EMAIL_RE = re.compile(r"[\w.+-]+@[\w.-]+\.\w+")


@dataclass
class EmailRow:
    sort_key: datetime
    date: str
    time: str
    direction: str
    from_addr: str
    to_addrs: str
    subject: str
    attachments: int
    party: str
    message_id: str


def _norm_email(addr: str) -> str:
    _, email = parseaddr(addr.strip())
    return email.lower()


def _extract_emails(text: str) -> list[str]:
    if not text:
        return []
    found = EMAIL_RE.findall(text.lower())
    return sorted(set(found))


def _party_label(email: str) -> str | None:
    e = email.lower()
    if e in EXACT_EMAILS:
        return e
    for label, domains in DOMAIN_RULES:
        if any(d in e for d in domains):
            return label
    return None


def _row_parties(from_addr: str, to_addrs: str, cc_addrs: str = "") -> list[str]:
    parties: list[str] = []
    for addr in _extract_emails(f"{from_addr} {to_addrs} {cc_addrs}"):
        if addr == MY_EMAIL.lower():
            continue
        label = _party_label(addr)
        if label and label not in parties:
            parties.append(label)
    return parties


def _build_gmail_query(months: int) -> str:
    end = datetime.now(IST).date()
    start = end - timedelta(days=months * 31)
    after = start.strftime("%Y/%m/%d")
    before = (end + timedelta(days=1)).strftime("%Y/%m/%d")
    clauses = [
        f"from:{addr} OR to:{addr}" for addr in sorted(EXACT_EMAILS)
    ]
    domain_clauses = [
        "from:homekraft.in OR to:homekraft.in",
        "from:pnbhousing.com OR to:pnbhousing.com",
        "from:atsgreen.com OR to:atsgreen.com",
        "from:atsgreens.com OR to:atsgreens.com",
        "from:hareraggm.gov.in OR to:hareraggm.gov.in",
    ]
    body = " OR ".join(clauses + domain_clauses)
    return f"after:{after} before:{before} ({body})"


def _unwrap_data(payload: object) -> dict:
    if isinstance(payload, str):
        try:
            payload = json.loads(payload)
        except json.JSONDecodeError:
            return {}
    if not isinstance(payload, dict):
        return {}
    if isinstance(payload.get("data"), dict):
        return payload["data"]
    if isinstance(payload.get("data"), str):
        try:
            inner = json.loads(payload["data"])
            return inner if isinstance(inner, dict) else payload
        except json.JSONDecodeError:
            return payload
    return payload


def _header_value(msg: dict, name: str) -> str:
    headers = msg.get("headers") or msg.get("payload", {}).get("headers") or []
    if isinstance(headers, dict):
        return str(headers.get(name, "") or headers.get(name.lower(), ""))
    for item in headers:
        if not isinstance(item, dict):
            continue
        key = (item.get("name") or item.get("key") or "").lower()
        if key == name.lower():
            return str(item.get("value", ""))
    return str(msg.get(name.lower(), "") or msg.get(name, ""))


def _attachment_count(msg: dict) -> int:
    for key in ("attachmentCount", "attachmentsCount", "numberOfAttachments"):
        val = msg.get(key)
        if isinstance(val, int):
            return val
    for key in ("attachmentList", "attachments", "attachment_list"):
        val = msg.get(key)
        if isinstance(val, list):
            return len(val)
    payload = msg.get("payload") or {}
    parts = payload.get("parts") or []
    count = 0

    def walk(parts_list: list) -> None:
        nonlocal count
        for part in parts_list:
            if not isinstance(part, dict):
                continue
            filename = part.get("filename") or ""
            body = part.get("body") or {}
            if filename and body.get("attachmentId"):
                count += 1
            nested = part.get("parts") or []
            if nested:
                walk(nested)

    walk(parts if isinstance(parts, list) else [])
    return count


def _internal_dt(msg: dict) -> datetime | None:
    raw = msg.get("internalDate") or msg.get("internal_date") or msg.get("date")
    if raw is None:
        return None
    if isinstance(raw, (int, float)):
        return datetime.fromtimestamp(float(raw) / 1000.0, tz=IST)
    if isinstance(raw, str) and raw.isdigit():
        return datetime.fromtimestamp(int(raw) / 1000.0, tz=IST)
    try:
        from email.utils import parsedate_to_datetime

        return parsedate_to_datetime(str(raw)).astimezone(IST)
    except (TypeError, ValueError, OverflowError):
        return None


def _direction(msg: dict, from_addr: str) -> str:
    labels = {str(x).upper() for x in (msg.get("labelIds") or msg.get("labels") or [])}
    if "SENT" in labels:
        return "Sent"
    if _norm_email(from_addr) == MY_EMAIL.lower():
        return "Sent"
    return "Received"


def _parse_message(msg: dict) -> EmailRow | None:
    message_id = str(
        msg.get("messageId")
        or msg.get("message_id")
        or msg.get("id")
        or ""
    )
    from_addr = _header_value(msg, "From") or str(msg.get("sender", ""))
    to_addrs = _header_value(msg, "To") or str(msg.get("recipient", ""))
    cc_addrs = _header_value(msg, "Cc")
    subject = _header_value(msg, "Subject") or str(msg.get("subject", ""))
    parties = _row_parties(from_addr, to_addrs, cc_addrs)
    if not parties:
        return None
    when = _internal_dt(msg)
    if when is None:
        return None
    return EmailRow(
        sort_key=when,
        date=when.strftime("%d.%m.%Y"),
        time=when.strftime("%H:%M"),
        direction=_direction(msg, from_addr),
        from_addr=from_addr,
        to_addrs=to_addrs,
        subject=subject.strip() or "(no subject)",
        attachments=_attachment_count(msg),
        party="; ".join(parties),
        message_id=message_id,
    )


def _decode_mime_header(value: str | None) -> str:
    if not value:
        return ""
    parts: list[str] = []
    for chunk, enc in decode_header(value):
        if isinstance(chunk, bytes):
            parts.append(chunk.decode(enc or "utf-8", errors="replace"))
        else:
            parts.append(str(chunk))
    return "".join(parts).strip()


def _load_local_env() -> None:
    """Load secrets from composio.env and/or gmail.env in project root."""
    for env_file in (LOCAL_ENV, GMAIL_ENV):
        if not env_file.exists():
            continue
        for line in env_file.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, value = line.split("=", 1)
            key, value = key.strip(), value.strip().strip('"').strip("'")
            if key and value and key not in os.environ:
                os.environ[key] = value


def _imap_attachment_count(msg: email.message.Message) -> int:
    count = 0
    for part in msg.walk():
        if part.get_content_maintype() == "multipart":
            continue
        disposition = (part.get_content_disposition() or "").lower()
        filename = part.get_filename()
        if disposition == "attachment" or (filename and disposition != "inline"):
            count += 1
    return count


def _parse_imap_message(
    msg: email.message.Message,
    uid: str,
    gmail_labels: list[str],
) -> EmailRow | None:
    from_addr = _decode_mime_header(msg.get("From", ""))
    to_addrs = _decode_mime_header(msg.get("To", ""))
    cc_addrs = _decode_mime_header(msg.get("Cc", ""))
    subject = _decode_mime_header(msg.get("Subject", "")) or "(no subject)"
    parties = _row_parties(from_addr, to_addrs, cc_addrs)
    if not parties:
        return None

    when: datetime | None = None
    try:
        when = parsedate_to_datetime(msg.get("Date", "")).astimezone(IST)
    except (TypeError, ValueError, OverflowError):
        pass
    if when is None:
        return None

    labels = {label.upper() for label in gmail_labels}
    if "SENT" in labels or _norm_email(from_addr) == MY_EMAIL.lower():
        direction = "Sent"
    else:
        direction = "Received"

    return EmailRow(
        sort_key=when,
        date=when.strftime("%d.%m.%Y"),
        time=when.strftime("%H:%M"),
        direction=direction,
        from_addr=from_addr,
        to_addrs=to_addrs,
        subject=subject,
        attachments=_imap_attachment_count(msg),
        party="; ".join(parties),
        message_id=msg.get("Message-ID", uid),
    )


def _gmail_labels_from_fetch(raw: bytes | tuple) -> list[str]:
    text = raw.decode("utf-8", errors="replace") if isinstance(raw, bytes) else str(raw)
    match = re.search(r"X-GM-LABELS \(([^)]*)\)", text)
    if not match:
        return []
    return re.findall(r'"([^"]+)"', match.group(1))


def _fetch_messages_imap(months: int) -> list[EmailRow]:
    user = os.environ.get("GMAIL_ADDRESS", MY_EMAIL)
    app_password = os.environ.get("GMAIL_APP_PASSWORD")
    if not app_password:
        raise SystemExit(
            "Gmail App Password required for direct inbox access.\n"
            "1. Google Account -> Security -> 2-Step Verification -> App passwords\n"
            "2. Create gmail.env in this folder:\n"
            "   GMAIL_ADDRESS=prof.vijyant.agarwal@gmail.com\n"
            "   GMAIL_APP_PASSWORD=your-16-char-app-password\n"
            "3. Run: python composio_gmail_email_sheet.py --source imap\n"
        )

    since = (datetime.now(IST) - timedelta(days=months * 31)).strftime("%d-%b-%Y")
    print(f"Connecting to Gmail IMAP ({user}) — live inbox, since {since}...")

    mail = imaplib.IMAP4_SSL("imap.gmail.com")
    mail.login(user, app_password.replace(" ", ""))

    folder = None
    for candidate in ('"[Gmail]/All Mail"', "INBOX"):
        try:
            status, _ = mail.select(candidate, readonly=True)
            if status == "OK":
                folder = candidate
                break
        except imaplib.IMAP4.error:
            continue
    if not folder:
        raise RuntimeError("Could not open Gmail folder (All Mail or INBOX).")

    print(f"Searching {folder} ...")
    status, data = mail.search(None, f"(SINCE {since})")
    if status != "OK":
        raise RuntimeError("Gmail search failed.")

    ids = data[0].split() if data and data[0] else []
    print(f"Scanning {len(ids)} messages from your mailbox...")

    rows: list[EmailRow] = []
    for idx, msg_id in enumerate(ids, 1):
        status, fetched = mail.fetch(msg_id, "(RFC822 X-GM-LABELS)")
        if status != "OK" or not fetched:
            continue
        raw_msg = b""
        labels: list[str] = []
        for item in fetched:
            if not isinstance(item, tuple) or len(item) < 2:
                continue
            meta, body = item[0], item[1]
            if isinstance(meta, bytes):
                labels = _gmail_labels_from_fetch(meta) or labels
            if isinstance(body, bytes) and body.startswith(b"Delivered-To:"):
                raw_msg = body
        if not raw_msg:
            continue
        msg = email.message_from_bytes(raw_msg)
        parsed = _parse_imap_message(msg, msg_id.decode(), labels)
        if parsed:
            rows.append(parsed)
        if idx % 100 == 0:
            print(f"  … processed {idx}/{len(ids)}")

    mail.logout()
    rows.sort(key=lambda r: r.sort_key, reverse=True)
    return rows


def _get_composio(api_key: str | None):
    from composio import Composio

    kwargs = {"toolkit_versions": {"gmail": "latest"}}
    if api_key:
        kwargs["api_key"] = api_key
    return Composio(**kwargs)


def _save_auth_config_id(auth_config_id: str) -> None:
    lines: list[str] = []
    if LOCAL_ENV.exists():
        for line in LOCAL_ENV.read_text(encoding="utf-8").splitlines():
            if line.strip().startswith("GMAIL_AUTH_CONFIG_ID="):
                continue
            lines.append(line.rstrip())
    else:
        lines = [f"COMPOSIO_API_KEY={os.environ.get('COMPOSIO_API_KEY', '')}"]
    lines.append(f"GMAIL_AUTH_CONFIG_ID={auth_config_id}")
    LOCAL_ENV.write_text("\n".join(lines).strip() + "\n", encoding="utf-8")
    os.environ["GMAIL_AUTH_CONFIG_ID"] = auth_config_id


def _wait_for_gmail_auth_config(composio, timeout_sec: int = 900) -> str:
    """Poll Composio until Gmail auth config appears (user creates it in dashboard)."""
    print(
        "\nWaiting for Gmail Auth Config in Composio dashboard "
        f"(up to {timeout_sec // 60} minutes)..."
    )
    print("Create: Auth Configs -> Add -> Gmail -> Composio managed OAuth -> Save\n")
    deadline = time.time() + timeout_sec
    while time.time() < deadline:
        configs = composio.client.auth_configs.list(toolkit_slug="gmail")
        items = getattr(configs, "items", None) or []
        if items:
            auth_id = items[0].id
            print(f"Found Gmail Auth Config: {auth_id}")
            _save_auth_config_id(auth_id)
            return auth_id
        time.sleep(5)
    raise SystemExit("Timed out waiting for Gmail Auth Config in Composio dashboard.")


def _open_connect_browser(redirect_url: str) -> None:
    print("\nOpen this URL to connect prof.vijyant.agarwal@gmail.com:\n")
    print(redirect_url)
    try:
        webbrowser.open(redirect_url)
    except OSError:
        pass
    try:
        import subprocess

        subprocess.Popen(["cmd", "/c", "start", "", redirect_url], shell=False)
    except OSError:
        pass


def _gmail_auth_config_id(composio, *, wait_setup: bool = False) -> str:
    env_id = os.environ.get("GMAIL_AUTH_CONFIG_ID", "").strip()
    if env_id:
        return env_id
    configs = composio.client.auth_configs.list(toolkit_slug="gmail")
    items = getattr(configs, "items", None) or []
    if items:
        auth_id = items[0].id
        _save_auth_config_id(auth_id)
        return auth_id
    if wait_setup:
        return _wait_for_gmail_auth_config(composio)
    raise SystemExit(
        "Gmail is not set up in Composio yet.\n\n"
        "Run with --wait-setup to open the dashboard and auto-connect:\n"
        "  python composio_gmail_email_sheet.py --source composio --wait-setup\n"
    )


def _ensure_gmail_connection(composio, connect: bool, *, wait_setup: bool = False) -> None:
    accounts = composio.connected_accounts.list(
        user_ids=[USER_ID],
        toolkit_slugs=["gmail"],
        statuses=["ACTIVE"],
    )
    if getattr(accounts, "items", None) and not connect:
        return

    auth_config_id = _gmail_auth_config_id(composio, wait_setup=wait_setup)
    req = composio.connected_accounts.link(USER_ID, auth_config_id)
    if getattr(req, "redirect_url", None):
        _open_connect_browser(req.redirect_url)
        print("\nWaiting for Gmail connection (up to 3 minutes)...")
        deadline = time.time() + 180
        while time.time() < deadline:
            try:
                req.wait_for_connection(timeout=10)
                print("Gmail connected to your live inbox.")
                return
            except Exception:
                time.sleep(3)
        raise SystemExit("Timed out waiting for Gmail connection. Re-run after authorizing.")


def _fetch_messages(composio, query: str) -> list[dict]:
    seen: set[str] = set()
    messages: list[dict] = []
    page_token: str | None = None

    while True:
        args = {
            "user_id": "me",
            "query": query,
            "max_results": 500,
            "verbose": True,
            "include_payload": True,
        }
        if page_token:
            args["page_token"] = page_token

        result = composio.tools.execute(
            "GMAIL_FETCH_EMAILS",
            arguments=args,
            user_id=USER_ID,
            dangerously_skip_version_check=True,
        )
        if not result.get("successful", True):
            raise RuntimeError(result.get("error") or "GMAIL_FETCH_EMAILS failed")

        data = _unwrap_data(result.get("data", result))
        batch = (
            data.get("messages")
            or data.get("emails")
            or data.get("items")
            or []
        )
        if isinstance(batch, dict):
            batch = batch.get("messages") or batch.get("items") or []

        for msg in batch:
            if not isinstance(msg, dict):
                continue
            mid = str(
                msg.get("messageId")
                or msg.get("message_id")
                or msg.get("id")
                or ""
            )
            if not mid or mid in seen:
                continue
            seen.add(mid)
            messages.append(msg)

        page_token = data.get("nextPageToken") or data.get("next_page_token")
        if not page_token:
            break

    return messages


def _write_markdown(rows: list[EmailRow], months: int, source: str) -> None:
    today = datetime.now(IST).strftime("%d %B %Y")
    lines = [
        "# Email Date & Time Sheet",
        "",
        f"**Mailbox:** {MY_EMAIL}  ",
        f"**Source:** live Gmail inbox ({source})  ",
        f"**Period:** last **{months} months**  ",
        f"**Parties:** ATS, DTCP, PNB Housing, HARERA and listed addresses  ",
        f"**Generated:** {today}  ",
        f"**Total emails:** {len(rows)}  ",
        "",
        "---",
        "",
        "| S.No. | Date | Time | Sent/Received | From | To | Subject | Attachments | Party |",
        "|-------|------|------|---------------|------|-----|---------|-------------|-------|",
    ]
    for i, row in enumerate(rows, 1):
        subj = row.subject.replace("|", "/")
        if len(subj) > 90:
            subj = subj[:87] + "..."
        lines.append(
            f"| {i} | {row.date} | {row.time} | {row.direction} | "
            f"{row.from_addr.replace('|', '/')} | {row.to_addrs.replace('|', '/')} | "
            f"{subj} | {row.attachments} | {row.party} |"
        )
    lines += [
        "",
        "---",
        "",
        f"*Source: live Gmail inbox via {source} · Regenerate: `python composio_gmail_email_sheet.py --source {source}`*",
    ]
    OUT_MD.write_text("\n".join(lines), encoding="utf-8")


def _write_xlsx(rows: list[EmailRow]) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Email Sheet"
    headers = [
        "S.No.",
        "Date",
        "Time",
        "Sent/Received",
        "From",
        "To",
        "Subject",
        "Attachments",
        "Party",
        "Message ID",
    ]
    header_fill = PatternFill("solid", fgColor="D9E2F3")
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = Font(bold=True)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for i, row in enumerate(rows, 2):
        ws.cell(row=i, column=1, value=i - 1)
        ws.cell(row=i, column=2, value=row.date)
        ws.cell(row=i, column=3, value=row.time)
        ws.cell(row=i, column=4, value=row.direction)
        ws.cell(row=i, column=5, value=row.from_addr)
        ws.cell(row=i, column=6, value=row.to_addrs)
        ws.cell(row=i, column=7, value=row.subject)
        ws.cell(row=i, column=8, value=row.attachments)
        ws.cell(row=i, column=9, value=row.party)
        ws.cell(row=i, column=10, value=row.message_id)

    widths = [6, 12, 8, 14, 34, 34, 48, 12, 24, 22]
    for idx, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(idx)].width = width
    ws.freeze_panes = "A2"
    wb.save(OUT_XLSX)


def _set_landscape(section) -> None:
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width, section.page_height = section.page_height, section.page_width


def _write_docx(rows: list[EmailRow], months: int) -> None:
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Inches(0.6)
    section.bottom_margin = Inches(0.6)
    section.left_margin = Inches(0.7)
    section.right_margin = Inches(0.7)
    _set_landscape(section)

    title = doc.add_heading("Email Date & Time Sheet", level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph(
        f"Mailbox: {MY_EMAIL} · Last {months} months · {len(rows)} emails · "
        f"Generated {datetime.now(IST).strftime('%d %B %Y')}"
    )

    headers = [
        "S.No.",
        "Date",
        "Time",
        "Sent/Recv",
        "From",
        "To",
        "Subject",
        "Att.",
        "Party",
    ]
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = "Table Grid"
    for col, header in enumerate(headers):
        cell = table.rows[0].cells[col]
        cell.text = header
        for run in cell.paragraphs[0].runs:
            run.bold = True
        shading = cell._tc.get_or_add_tcPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:fill"), "D9E2F3")
        shading.append(shd)

    for i, row in enumerate(rows, 1):
        vals = [
            str(i),
            row.date,
            row.time,
            row.direction,
            row.from_addr,
            row.to_addrs,
            row.subject[:120],
            str(row.attachments),
            row.party,
        ]
        for col, val in enumerate(vals):
            table.rows[i].cells[col].text = val

    doc.save(OUT_DOCX)


def _resolve_source(requested: str) -> str:
    if requested != "auto":
        return requested
    if os.environ.get("GMAIL_APP_PASSWORD"):
        return "imap"
    if os.environ.get("COMPOSIO_API_KEY"):
        return "composio"
    _print_setup_help()
    raise SystemExit(1)


def _print_setup_help() -> None:
    print(
        "Cannot access your live Gmail inbox yet. Choose one setup:\n\n"
        "OPTION A — Direct Gmail (recommended, reads your inbox via IMAP):\n"
        "  1. Google Account -> Security -> 2-Step Verification -> App passwords\n"
        "  2. Create gmail.env in this folder:\n"
        "       GMAIL_ADDRESS=prof.vijyant.agarwal@gmail.com\n"
        "       GMAIL_APP_PASSWORD=xxxx xxxx xxxx xxxx\n"
        "  3. Run: python composio_gmail_email_sheet.py --source imap\n\n"
        "OPTION B — Composio:\n"
        "  1. https://platform.composio.dev -> API key\n"
        "  2. Create composio.env: COMPOSIO_API_KEY=your-key\n"
        "  3. Run: python composio_gmail_email_sheet.py --source composio --connect\n",
        file=sys.stderr,
    )


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Email date/time sheet from your LIVE Gmail inbox"
    )
    parser.add_argument("--api-key", help="Composio API key (or set COMPOSIO_API_KEY)")
    parser.add_argument("--months", type=int, default=2, help="Months of email history")
    parser.add_argument(
        "--source",
        choices=["auto", "composio", "imap"],
        default="auto",
        help="auto = Gmail IMAP if gmail.env exists, else Composio",
    )
    parser.add_argument(
        "--connect",
        action="store_true",
        help="Force Gmail OAuth connection flow (Composio only)",
    )
    parser.add_argument(
        "--wait-setup",
        action="store_true",
        help="Open Composio dashboard, wait for Gmail Auth Config, then connect inbox",
    )
    args = parser.parse_args()
    _load_local_env()
    source = _resolve_source(args.source)

    if source == "imap":
        rows = _fetch_messages_imap(args.months)
    else:
        try:
            composio = _get_composio(args.api_key)
        except Exception as exc:
            if args.source == "composio":
                print(
                    "Composio API key missing.\n"
                    "Create composio.env in this folder:\n"
                    "  COMPOSIO_API_KEY=your-key-from-platform.composio.dev\n"
                    "Then re-run:\n"
                    "  python composio_gmail_email_sheet.py --source composio --connect\n",
                    file=sys.stderr,
                )
            else:
                _print_setup_help()
            raise SystemExit(1) from exc

        if args.wait_setup:
            for url in (
                "https://platform.composio.dev/auth-configs",
                "https://app.composio.dev/auth-configs",
            ):
                try:
                    webbrowser.open(url)
                except OSError:
                    pass
            print("Opened Composio Auth Configs in your browser.")
            print("Create Gmail auth config (Composio managed OAuth), then wait...")

        _ensure_gmail_connection(
            composio,
            connect=args.connect or args.wait_setup,
            wait_setup=args.wait_setup,
        )
        query = _build_gmail_query(args.months)
        print(f"Fetching live Gmail via Composio (last {args.months} months)...")
        print(f"Query: {query[:160]}...")

        raw_messages = _fetch_messages(composio, query)
        rows = []
        for msg in raw_messages:
            parsed = _parse_message(msg)
            if parsed:
                rows.append(parsed)
        rows.sort(key=lambda r: r.sort_key, reverse=True)

    if not rows:
        print("No matching emails found in your inbox for the selected parties/date range.")

    _write_markdown(rows, args.months, source)
    _write_xlsx(rows)
    _write_docx(rows, args.months)

    print(f"Written {OUT_MD.name} — {len(rows)} emails from your live inbox")
    print(f"Written {OUT_XLSX.name}")
    print(f"Written {OUT_DOCX.name}")


if __name__ == "__main__":
    main()
