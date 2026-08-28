"""Convert 112_Gmail_22062026_ATS_Meeting_PreEMI_Invite.pdf to markdown."""
from pathlib import Path

from convert_gmail_email_common import convert_pdf

PNB_ROOT = Path(__file__).resolve().parents[1]
PDF = PNB_ROOT / "attachments" / "112_Gmail_22062026_ATS_Meeting_PreEMI_Invite.pdf"


def main() -> None:
    if not PDF.is_file():
        raise SystemExit(f"PDF not found: {PDF}")
    out = convert_pdf(PDF)
    print(f"Source: {PDF}")
    print(f"Output: {out} ({out.stat().st_size} bytes)")


if __name__ == "__main__":
    main()
