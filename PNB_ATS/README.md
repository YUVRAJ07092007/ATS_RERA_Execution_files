# PNB_ATS — folder layout

Pre-EMI / PNB Housing correspondence and supporting documents for **Unit 3051**, ATS Grandstand.

## Structure

| Folder | Purpose |
|--------|---------|
| **email_letters/** | Draft outgoing emails and letters (.md source + .docx for sending). Edit .md first; regenerate .docx as needed. |
| **attachments/** | **All PDFs** for this workstream — numbered **101–114** (chronological; higher number = later document). Use for sending enclosures, background reference, and markdown conversion sources. |
| **md_pnb_ats/** | Markdown conversions of source PDFs, convert_*.py scripts, organize_pnb_pdfs.py, and rename_pnb_pdfs.py (run from repo with PyMuPDF installed). |

## Typical send workflow

1. Finalize draft in **email_letters/*.md** (counsel review).
2. Export or copy to **.docx** in the same folder for Gmail / Word.
3. For the **PNB procedure letter**, attach PDFs **111, 105, 101, 109** from **attachments/** (see PNB letter Encl. mapping below).
4. After sending the PNB letter, save Gmail export as **114_Gmail_23062026_REQUEST_FOR_DIRECT_EMI_PAYMENT_BY_BUILDER.pdf** (or regenerate and replace).

## attachments/ — numbered index

Files use a **101–114** prefix with descriptive names: `{slot}_{DocType}_{DDMMYYYY}_{TITLE}.pdf`.

| # | Filename | Document | PNB letter Encl. |
|---|----------|----------|------------------|
| 101 | 101_HARERA_ORDER_DATED_08082024.pdf | HARERA order — s.39 rectification (08.08.2024) | **Encl.3** |
| 102 | 102_Tripartite_AGREEMENT_07052018.pdf | Tripartite agreement (07.05.2018) | — |
| 103 | 103_PNB_Sanction_Letter_19092018.pdf | PNB Housing sanction letter (19.09.2018) | — |
| 104 | 104_BBA_Agreement_to_Sell_18102018.pdf | BBA / Agreement to Sell (18.10.2018) | — |
| 105 | 105_ATS_Builder_preEMI_Letter_23102018.pdf | Builder pre-EMI / subvention letter (23.10.2018) | **Encl.2** |
| 106 | 106_PNB_Loan_Account_Details.pdf | Loan account details / summary (Nov 2021) | — |
| 107 | 107_HARERA_ORDER_DATED_18042024.pdf | HARERA order (18.04.2024) — background reference only | — |
| 108 | 108_Allottee_Letter_PNB_PreEMI_Reimbursement_02012025.pdf | Allottee letter to PNB re pre-EMI reimbursement (02.01.2025) | — |
| 109 | 109_Gmail_18062026_REQUEST_Loan_Account_HRERA_Compliance.pdf | Promoter email requesting loan details (18.06.2026) | **Encl.4** |
| 110 | 110_PNB_Full_Loan_Statement_22062026.pdf | Full loan account statement (22.06.2026) | — |
| 111 | 111_PNB_Statement_First_Page_Subvention_Yes_22062026.pdf | PNBHFL statement first page — Subvention: Yes (22.06.2026) | **Encl.1** |
| 112 | 112_Gmail_22062026_ATS_Meeting_PreEMI_Invite.pdf | ATS Gmail — meeting invite / pre-EMI docs (22.06.2026) | — |
| 113 | 113_Gmail_22062026_ATS_GM251D_Meeting_PreEMI.pdf | ATS Gmail — DTCP grievance meeting (22.06.2026; image export) | — |
| 114 | 114_Gmail_23062026_REQUEST_FOR_DIRECT_EMI_PAYMENT_BY_BUILDER.pdf | Gmail export of sent PNB procedure email (23.06.2026) | — |

### Legacy filename mapping (100-series / Encl.* / old 01–14 → attachments/)

Older files used a **100–107** prefix, **Encl.*** names, or mistaken **01–14** prefixes. They map to the canonical **attachments/** index below. Run `python PNB_ATS/md_pnb_ats/organize_pnb_pdfs.py` to move strays, delete byte-identical duplicates, and auto-correct mistaken **01–14** renames back to **101–114**.

| Legacy name | → # | Canonical file |
|-------------|-----|----------------|
| Encl.3_Section39_rectification_08082024.pdf | **101** | 101_HARERA_ORDER_DATED_08082024.pdf |
| 107_HARERA_order_08082024.pdf | **101** | 101_HARERA_ORDER_DATED_08082024.pdf |
| 100_tripartite _AGREEMENT _07052018.pdf | **102** | 102_Tripartite_AGREEMENT_07052018.pdf |
| Encl.1_Sanction_letter_19092018.pdf | **103** | 103_PNB_Sanction_Letter_19092018.pdf |
| BBA_Agreement_to_Sell_18102018.pdf | **104** | 104_BBA_Agreement_to_Sell_18102018.pdf |
| Encl.2_Builder_preEMI_letter_23102018.pdf | **105** | 105_ATS_Builder_preEMI_Letter_23102018.pdf |
| 102_loan_account_details_PNB Housing.pdf | **106** | 106_PNB_Loan_Account_Details.pdf |
| Encl.3_HARERA_order_18042024.pdf | **107** | 107_HARERA_ORDER_DATED_18042024.pdf |
| 105_letter_to_PNB_Pre-EMI Reimbursement_pdf.pdf | **108** | 108_Allottee_Letter_PNB_PreEMI_Reimbursement_02012025.pdf |
| 107_Gmail_Request_for_SharingLoan Account_Details...pdf | **109** | 109_Gmail_18062026_REQUEST_Loan_Account_HRERA_Compliance.pdf |
| 22062026_loan_Account_Statement.pdf | **110** | 110_PNB_Full_Loan_Statement_22062026.pdf |
| (statement first page) | **111** | 111_PNB_Statement_First_Page_Subvention_Yes_22062026.pdf |
| Encl.4_Builder_email_22062026.pdf | **112** | 112_Gmail_22062026_ATS_Meeting_PreEMI_Invite.pdf |
| 22062026_gmail_from_ATS_GM251D_1.pdf | **113** | 113_Gmail_22062026_ATS_GM251D_Meeting_PreEMI.pdf |
| Gmail — PNB procedure email export | **114** | 114_Gmail_23062026_REQUEST_FOR_DIRECT_EMI_PAYMENT_BY_BUILDER.pdf |

### md_pnb_ats/ — markdown conversions

Markdown files use the same **101–114** prefix as their source PDF in **attachments/**:

| # | Markdown | Convert script |
|---|----------|----------------|
| 102 | 102_Tripartite_AGREEMENT_07052018.md | — (manual/OCR) |
| 103 | 103_PNB_Sanction_Letter_19092018.md | convert_103_sanction_letter.py |
| 105 | 105_ATS_Builder_preEMI_Letter_23102018.md | convert_105_builder_preEMI_letter.py |
| 106 | 106_PNB_Loan_Account_Details.md | convert_106_loan_account.py |
| 110 | 110_PNB_Full_Loan_Statement_22062026.md | convert_110_full_loan_statement.py |
| 111 | 111_PNB_Statement_First_Page_Subvention_Yes_22062026.md | convert_111_statement_first_page.py |
| 112 | 112_Gmail_22062026_ATS_Meeting_PreEMI_Invite.md | convert_112_ATS_email_meeting_invite.py |
| 113 | 113_Gmail_22062026_ATS_GM251D_Meeting_PreEMI.md | convert_113_Gmail_from_ATS_GM251D_1.py |

Legacy md names (e.g. `01_tripartite_*`, `03_Sanction_*`, `22062026_*`, `3_ATS_Subvention_*`) were renamed to match the **101–114** PDF index above.

### PNB letter — attach these four

When sending **23062026_PNB_PreEMI_Procedure_Unit_3051**, attach:

| Letter Encl. | Attach # | File |
|--------------|----------|------|
| Encl.1 | **111** | 111_PNB_Statement_First_Page_Subvention_Yes_22062026.pdf |
| Encl.2 | **105** | 105_ATS_Builder_preEMI_Letter_23102018.pdf |
| Encl.3 | **101** | 101_HARERA_ORDER_DATED_08082024.pdf |
| Encl.4 | **109** | 109_Gmail_18062026_REQUEST_Loan_Account_HRERA_Compliance.pdf |

Do **not** attach 102–104, 106–108, 110, 112–113, or 114 unless counsel advises. **107** (18.04.2024 order) is cited in the letter text only.

## Notes

- **Slot 101:** HARERA order dated **08.08.2024** (Encl.3). Pattern: `101_HARERA_ORDER_DATED_08082024.pdf`.
- **Slot 114:** Gmail export of the **23.06.2026** PNB procedure letter after sending — `114_Gmail_23062026_REQUEST_FOR_DIRECT_EMI_PAYMENT_BY_BUILDER.pdf`.
- **23062026_ATS_IFSC_PreEMI_Unit_3051_FINAL.md** is marked **FINAL — do not edit**. Longer counsel draft: **23062026_Reply_to_ATS_IFSC_and_PreEMI_Procedure_Unit_3051.md** (superseded).
- Root of **PNB_ATS/** holds only this README and the three subfolders above. **All PDFs live in attachments/ only** — no duplicates elsewhere.
