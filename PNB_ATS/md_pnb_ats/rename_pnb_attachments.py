"""Restore attachments/101-114 content to correct slots and rename descriptively.

Run from repo root:
    python PNB_ATS/md_pnb_ats/rename_pnb_attachments.py

Content verified 2026-06-23 via pypdf first-page extraction + MD5 hash mapping.
"""
from __future__ import annotations

import hashlib
from pathlib import Path

ATT = Path(__file__).resolve().parents[1] / "attachments"

TARGET_BY_HASH: dict[str, tuple[int, str]] = {
    "b89c46703de6": (114, "114_Gmail_23062026_REQUEST_FOR_DIRECT_EMI_PAYMENT_BY_BUILDER.pdf"),
    "872816aca9ee": (101, "101_HARERA_ORDER_DATED_08082024.pdf"),
    "e099627a434f": (102, "102_Tripartite_AGREEMENT_07052018.pdf"),
    "a314d01d6112": (103, "103_PNB_Sanction_Letter_19092018.pdf"),
    "11e93df3caa7": (104, "104_BBA_Agreement_to_Sell_18102018.pdf"),
    "4dd24dde5a56": (105, "105_ATS_Builder_preEMI_Letter_23102018.pdf"),
    "2556b05b3b5c": (106, "106_PNB_Loan_Account_Details.pdf"),
    "162ef2613d35": (107, "107_HARERA_ORDER_DATED_18042024.pdf"),
    "66922019a757": (108, "108_Allottee_Letter_PNB_PreEMI_Reimbursement_02012025.pdf"),
    "32efbd1fe57c": (109, "109_Gmail_18062026_REQUEST_Loan_Account_HRERA_Compliance.pdf"),
    "b2a83856a79e": (110, "110_PNB_Full_Loan_Statement_22062026.pdf"),
    "f809fe3df805": (111, "111_PNB_Statement_First_Page_Subvention_Yes_22062026.pdf"),
    "acb92833bac5": (112, "112_Gmail_22062026_ATS_Meeting_PreEMI_Invite.pdf"),
    "b83217eb9b3a": (113, "113_Gmail_22062026_ATS_GM251D_Meeting_PreEMI.pdf"),
}


def md5(path: Path) -> str:
    h = hashlib.md5()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def main() -> None:
    pdfs = sorted(ATT.glob("*.pdf"))
    if len(pdfs) != 14:
        raise SystemExit(f"Expected 14 PDFs in attachments/, found {len(pdfs)}")

    plan: list[tuple[Path, Path]] = []
    seen_hashes: set[str] = set()

    for src in pdfs:
        h = md5(src)
        key = h[:12]
        if key not in TARGET_BY_HASH:
            raise SystemExit(f"Unknown PDF content: {src.name} md5={h}")
        if key in seen_hashes:
            raise SystemExit(f"Duplicate content hash {key} — manual review needed")
        seen_hashes.add(key)
        _slot, dest_name = TARGET_BY_HASH[key]
        plan.append((src, ATT / dest_name))

    dest_names = [d.name for _, d in plan]
    if len(dest_names) != len(set(dest_names)):
        raise SystemExit("Destination name collision in plan")

    # Phase 1: move all to temp names
    temps: list[tuple[Path, Path]] = []
    for i, (src, dest) in enumerate(plan):
        tmp = ATT / f"__renaming_{i:02d}.pdf"
        if src.resolve() == dest.resolve():
            temps.append((src, dest))
            continue
        src.rename(tmp)
        temps.append((tmp, dest))

    # Phase 2: temp -> final
    for tmp, dest in temps:
        if tmp.resolve() == dest.resolve():
            continue
        if dest.exists():
            dest.unlink()
        tmp.rename(dest)

    print("Renamed / permuted 14 attachment PDFs:")
    for _src, dest in sorted(plan, key=lambda x: x[1].name):
        print(f"  -> {dest.name}")


if __name__ == "__main__":
    main()
