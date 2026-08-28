"""Move stray PNB PDFs into attachments/; dedupe against 101-114 index.

Run from repo root:
    python PNB_ATS/md_pnb_ats/organize_pnb_pdfs.py

Legacy 100-series, Encl.*, and old 01-14 names map to canonical 101-114 files.
"""
from __future__ import annotations

import hashlib
import re
import shutil
from pathlib import Path

PNB_ROOT = Path(__file__).resolve().parents[1]
ATTACH = PNB_ROOT / "attachments"
SKIP_DIRS = {"attachments", "email_letters", "md_pnb_ats", "archive", "reference_pdfs"}

CANONICAL_BY_SLOT: dict[int, str] = {
    1: "101_HARERA_ORDER_DATED_08082024.pdf",
    2: "102_Tripartite_AGREEMENT_07052018.pdf",
    3: "103_PNB_Sanction_Letter_19092018.pdf",
    4: "104_BBA_Agreement_to_Sell_18102018.pdf",
    5: "105_ATS_Builder_preEMI_Letter_23102018.pdf",
    6: "106_PNB_Loan_Account_Details.pdf",
    7: "107_HARERA_ORDER_DATED_18042024.pdf",
    8: "108_Allottee_Letter_PNB_PreEMI_Reimbursement_02012025.pdf",
    9: "109_Gmail_18062026_REQUEST_Loan_Account_HRERA_Compliance.pdf",
    10: "110_PNB_Full_Loan_Statement_22062026.pdf",
    11: "111_PNB_Statement_First_Page_Subvention_Yes_22062026.pdf",
    12: "112_Gmail_22062026_ATS_Meeting_PreEMI_Invite.pdf",
    13: "113_Gmail_22062026_ATS_GM251D_Meeting_PreEMI.pdf",
    14: "114_Gmail_23062026_REQUEST_FOR_DIRECT_EMI_PAYMENT_BY_BUILDER.pdf",
}

CANONICAL_BY_NUM: dict[int, str] = {100 + k: v for k, v in CANONICAL_BY_SLOT.items()}

LEGACY_MAP: dict[str, str] = {
    "100_tripartite _agreement _07052018": CANONICAL_BY_SLOT[2],
    "100_tripartite_agreement_07052018": CANONICAL_BY_SLOT[2],
    "102_loan_account_details_pnb housing": CANONICAL_BY_SLOT[6],
    "102_loan_account_details_pnb_housing": CANONICAL_BY_SLOT[6],
    "105_letter_to_pnb_pre-emi reimbursement_pdf": CANONICAL_BY_SLOT[8],
    "105_letter_to_pnb_preemi_reimbursement_pdf": CANONICAL_BY_SLOT[8],
    "107_gmail_request_for_sharingloan account_details for compliance with hrera order": CANONICAL_BY_SLOT[9],
    "107_gmail_request_loan_account_hrera_compliance": CANONICAL_BY_SLOT[9],
    "109_gmail_request_loan_account_hrera_compliance": CANONICAL_BY_SLOT[9],
    "22062026_loan_account_statement": CANONICAL_BY_SLOT[10],
    "bba_agreement_to_sell_18102018": CANONICAL_BY_SLOT[4],
    "encl.1_sanction_letter_19092018": CANONICAL_BY_SLOT[3],
    "encl.2_builder_preemi_letter_23102018": CANONICAL_BY_SLOT[5],
    "encl.3_harera_order_18042024": CANONICAL_BY_SLOT[7],
    "encl.3_section39_rectification_08082024": CANONICAL_BY_SLOT[1],
    "107_harera_order_08082024": CANONICAL_BY_SLOT[1],
    "101_harera_order_dated_08082024": CANONICAL_BY_SLOT[1],
    "encl.4_builder_email_22062026": CANONICAL_BY_SLOT[12],
    "gmail - request for sharing loan account details for compliance with hrera order": CANONICAL_BY_SLOT[9],
    "pnb sanction letter - vijyant agarwal (2)": CANONICAL_BY_SLOT[3],
    "pnb statement_vijyant_21112024131404 (1)": CANONICAL_BY_SLOT[6],
    "order  april 2024  rera ats ": CANONICAL_BY_SLOT[7],
    "letter ": CANONICAL_BY_SLOT[8],
    "114_gmail_pnb_procedure_email_export": CANONICAL_BY_SLOT[14],
    "gmail - loan a_c hou_gur_0918_578047 request for procedure and document checklist": CANONICAL_BY_SLOT[14],
    # Old descriptive names (pre-rename)
    "102_tripartite_agreement_07052018": CANONICAL_BY_SLOT[2],
    "103_sanction_letter_19092018": CANONICAL_BY_SLOT[3],
    "105_builder_preemi_letter_23102018": CANONICAL_BY_SLOT[5],
    "106_PNB_Loan_Account_Details": CANONICAL_BY_SLOT[6],
    "107_harera_order_18042024": CANONICAL_BY_SLOT[7],
    "108_letter_to_pnb_preemi_reimbursement_02012025": CANONICAL_BY_SLOT[8],
    "110_PNB_Full_Loan_Statement_22062026": CANONICAL_BY_SLOT[10],
    "111_statement_of_account_first_page_22062026": CANONICAL_BY_SLOT[11],
    "112_ats_email_meeting_invite_22062026": CANONICAL_BY_SLOT[12],
    "113_gmail_from_ats_gm251d_1": CANONICAL_BY_SLOT[13],
}

OLD_01_14_MAP: dict[str, str] = {
    "01_tripartite_agreement_07052018": CANONICAL_BY_SLOT[2],
    "02_bba_agreement_to_sell_18102018": CANONICAL_BY_SLOT[4],
    "03_sanction_letter_19092018": CANONICAL_BY_SLOT[3],
    "04_loan_account_details": CANONICAL_BY_SLOT[6],
    "05_letter_to_pnb_preemi_reimbursement_02012025": CANONICAL_BY_SLOT[8],
    "06_full_loan_statement_22062026": CANONICAL_BY_SLOT[10],
    "07_builder_preemi_letter_23102018": CANONICAL_BY_SLOT[5],
    "08_harera_order_18042024": CANONICAL_BY_SLOT[7],
    "09_harera_order_08082024": CANONICAL_BY_SLOT[1],
    "10_ats_email_meeting_invite_22062026": CANONICAL_BY_SLOT[12],
    "11_gmail_from_ats_gm251d_1": CANONICAL_BY_SLOT[13],
    "12_gmail_request_loan_account_hrera_compliance": CANONICAL_BY_SLOT[9],
    "13_statement_of_account_first_page_22062026": CANONICAL_BY_SLOT[11],
    "14_gmail_pnb_procedure_email_export": CANONICAL_BY_SLOT[14],
}

NUM_PREFIX_RE = re.compile(r"^(\d{2,3})_", re.IGNORECASE)
CANONICAL_NUM_RE = re.compile(r"^10[1-9]_|^11[0-4]_", re.IGNORECASE)


def norm_stem(name: str) -> str:
    stem = Path(name).stem.lower()
    return re.sub(r"\s+", " ", stem).strip()


def md5(path: Path) -> str:
    h = hashlib.md5()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def attachment_hashes() -> dict[str, str]:
    out: dict[str, str] = {}
    if not ATTACH.is_dir():
        return out
    for p in ATTACH.glob("*.pdf"):
        if CANONICAL_NUM_RE.match(p.name):
            out[p.name] = md5(p)
    return out


def next_attachment_number(existing: set[int]) -> int:
    base = max(existing, default=100)
    n = base + 1
    while n in existing:
        n += 1
    return n


def resolve_target(name: str, existing_names: set[str]) -> str | None:
    key = norm_stem(name)
    if key in LEGACY_MAP:
        return LEGACY_MAP[key]
    if key in OLD_01_14_MAP:
        return OLD_01_14_MAP[key]
    for v in CANONICAL_BY_SLOT.values():
        if norm_stem(v) == key:
            return v
    m = NUM_PREFIX_RE.match(name)
    if not m:
        return None
    num = int(m.group(1))
    if 101 <= num <= 114 and num in CANONICAL_BY_NUM:
        return CANONICAL_BY_NUM[num]
    if 1 <= num <= 14 and num in CANONICAL_BY_SLOT:
        return CANONICAL_BY_SLOT[num]
    return None


def fix_legacy_01_14_in_attachments(hashes: dict[str, str]) -> list[str]:
    """Rename mistaken 01-14 files in attachments/ to canonical 101-114."""
    renamed: list[str] = []
    pending: list[tuple[Path, Path]] = []
    for p in sorted(ATTACH.glob("*.pdf")):
        if CANONICAL_NUM_RE.match(p.name):
            continue
        target_name = OLD_01_14_MAP.get(norm_stem(p.name))
        if not target_name or p.name == target_name:
            continue
        pending.append((p, ATTACH / ("__tmp__" + target_name)))
    for src, tmp in pending:
        src.rename(tmp)
    for src, tmp in pending:
        dest = ATTACH / tmp.name.replace("__tmp__", "", 1)
        tmp.rename(dest)
        if src.name in hashes:
            del hashes[src.name]
        hashes[dest.name] = md5(dest)
        renamed.append(f"attachments/{src.name} -> {dest.name}")
    return renamed


def iter_stray_pdfs() -> list[Path]:
    found: list[Path] = []
    if not PNB_ROOT.is_dir():
        return found
    for p in PNB_ROOT.rglob("*.pdf"):
        rel = p.relative_to(PNB_ROOT)
        if rel.parts[0] in SKIP_DIRS:
            continue
        if rel.parts[0] == "attachments":
            if len(rel.parts) > 1 and rel.parts[1] == "_staging_merge":
                found.append(p)
            continue
        found.append(p)
    return sorted(found)


def main() -> None:
    ATTACH.mkdir(parents=True, exist_ok=True)
    hashes = attachment_hashes()
    existing_names = set(hashes)
    existing_nums = {
        int(n.split("_", 1)[0])
        for n in existing_names
        if CANONICAL_NUM_RE.match(n)
    }

    renamed = fix_legacy_01_14_in_attachments(hashes)
    existing_names = set(hashes)
    existing_nums = {
        int(n.split("_", 1)[0])
        for n in existing_names
        if CANONICAL_NUM_RE.match(n)
    }

    moved: list[str] = []
    deduped: list[str] = []
    skipped: list[str] = []
    added: list[str] = []

    for src in iter_stray_pdfs():
        rel = src.relative_to(PNB_ROOT)
        target_name = resolve_target(src.name, existing_names)
        src_hash = md5(src)

        if target_name and target_name in hashes:
            if src_hash == hashes[target_name]:
                src.unlink()
                deduped.append(f"{rel} (duplicate of attachments/{target_name})")
            else:
                skipped.append(
                    f"{rel} (conflicts with attachments/{target_name}; different content)"
                )
            continue

        if target_name:
            dest = ATTACH / target_name
        else:
            n = next_attachment_number(existing_nums)
            dest = ATTACH / f"{n}_{src.stem}.pdf"
            existing_nums.add(n)
            target_name = dest.name
            added.append(f"{rel} -> attachments/{target_name} (new slot)")

        if dest.exists() and md5(dest) == src_hash:
            src.unlink()
            deduped.append(f"{rel} (duplicate of attachments/{dest.name})")
            continue

        shutil.move(str(src), str(dest))
        hashes[target_name] = src_hash
        existing_names.add(target_name)
        moved.append(f"{rel} -> attachments/{target_name}")

    staging = ATTACH / "_staging_merge"
    if staging.is_dir():
        for p in list(staging.glob("*.pdf")):
            if p.name in hashes and md5(p) == hashes[p.name]:
                p.unlink()
                deduped.append(f"attachments/_staging_merge/{p.name} (staging duplicate)")
        if staging.is_dir() and not any(staging.iterdir()):
            staging.rmdir()

    print("=== PNB PDF organize ===")
    print(f"Attachments index: {len(hashes)} files in {ATTACH}")
    if renamed:
        print("\nRenamed (01-14 -> 101-114):")
        for line in renamed:
            print(f"  {line}")
    if moved:
        print("\nMoved:")
        for line in moved:
            print(f"  {line}")
    if deduped:
        print("\nDeduplicated (deleted stray copy):")
        for line in deduped:
            print(f"  {line}")
    if added:
        print("\nNew attachment slots:")
        for line in added:
            print(f"  {line}")
    if skipped:
        print("\nSkipped (manual review):")
        for line in skipped:
            print(f"  {line}")
    if not any([renamed, moved, deduped, added, skipped]):
        print("\nNo stray PDFs found - attachments/ is already canonical (101-114).")


if __name__ == "__main__":
    main()
