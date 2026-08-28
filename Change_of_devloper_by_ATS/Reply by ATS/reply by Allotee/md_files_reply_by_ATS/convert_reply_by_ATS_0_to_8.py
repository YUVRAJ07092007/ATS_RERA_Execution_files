"""Convert Reply by ATS files 0-8 to markdown in this folder."""
from __future__ import annotations

import re
import subprocess
from pathlib import Path

try:
    import fitz
except ImportError:
    subprocess.check_call(["pip", "install", "pymupdf", "-q"])
    import fitz

SRC = Path(__file__).resolve().parents[2]  # Reply by ATS
OUT = Path(__file__).resolve().parent

FILE0_FALLBACK = SRC / (
    "Gmail - Representation _ Objection Regarding Non-Service of Individual Notice "
    "and Non-Disclosure of Revised Sanctioned Plans of the project GRANDSTAND "
    "sector 99A ,Gurugram (1).pdf"
)

FILES = sorted(
    p for p in SRC.iterdir()
    if p.is_file() and re.match(r"^[0-8]", p.name)
)


def slug_stem(path: Path) -> str:
    return path.stem.replace(" ", "_").replace("/", "-")


def header(title: str, source: str) -> str:
    return f"# {title}\n\nConverted from: `{source}`\n\n---\n\n"


def convert_docx(path: Path) -> str:
    result = subprocess.run(
        ["pandoc", str(path), "-t", "markdown", "--wrap=none"],
        capture_output=True,
        text=True,
        check=True,
    )
    body = result.stdout.strip()
    title = path.stem.lstrip("0123456789-_").strip() or path.stem
    return header(title, path.name) + body + "\n"


def convert_pdf(path: Path) -> str:
    doc = fitz.open(path)
    pages: list[str] = []
    for i, page in enumerate(doc, start=1):
        text = page.get_text().strip()
        if text:
            pages.append(f"## Page {i}\n\n```text\n{text}\n```")
        else:
            pages.append(f"## Page {i}\n\n*(No extractable text — likely image/drawing.)*")
    title = path.stem.lstrip("0123456789-_").strip() or path.stem
    body = "\n\n".join(pages) if pages else "*(No text extracted.)*"
    return header(title, path.name) + body + "\n"


def convert_image(path: Path) -> str:
    title = path.stem.lstrip("0123456789-_").strip() or path.stem
    rel = f"../{path.name}"
    return (
        header(title, path.name)
        + f"![{path.name}]({rel})\n\n"
        + f"*Source image: `{path.name}` in `Reply by ATS` folder.*\n"
    )


def convert_file(path: Path) -> Path:
    ext = path.suffix.lower()
    if ext == ".docx":
        content = convert_docx(path)
    elif ext == ".pdf":
        content = convert_pdf(path)
    elif ext in {".jpg", ".jpeg", ".png"}:
        content = convert_image(path)
    else:
        raise ValueError(f"Unsupported type: {path}")

    out_name = f"{path.stem}.md"
    out_path = OUT / out_name
    out_path.write_text(content, encoding="utf-8")
    return out_path


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    for path in FILES:
        try:
            out = convert_file(path)
            print(f"OK  {path.name} -> {out.name}")
        except OSError as exc:
            print(f"SKIP  {path.name} ({exc})")

    file0_name = (
        "0_Gmail - Representation _ Objection Regarding Non-Service of Individual Notice "
        "and Non-Disclosure of Revised Sanctioned Plans of the project GRANDSTAND "
        "sector 99A ,Gurugram.pdf"
    )
    file0_out = OUT / (
        "0_Gmail - Representation _ Objection Regarding Non-Service of Individual Notice "
        "and Non-Disclosure of Revised Sanctioned Plans of the project GRANDSTAND "
        "sector 99A ,Gurugram.md"
    )
    if not file0_out.exists() and FILE0_FALLBACK.exists():
        content = convert_pdf(FILE0_FALLBACK)
        note = (
            "\n\n*Note: Original `0_Gmail...pdf` in `Reply by ATS` was unreadable; "
            f"converted from `{FILE0_FALLBACK.name}`.*\n"
        )
        file0_out.write_text(content.rstrip() + note, encoding="utf-8")
        print(f"OK  {FILE0_FALLBACK.name} -> {file0_out.name} (file 0 fallback)")


if __name__ == "__main__":
    main()
