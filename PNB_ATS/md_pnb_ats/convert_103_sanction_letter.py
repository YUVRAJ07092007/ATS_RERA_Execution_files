"""Convert 103_PNB_Sanction_Letter_19092018.pdf to markdown."""
from pathlib import Path

import fitz

REF = Path(__file__).resolve().parents[1] / "attachments"
PDF = REF / "103_PNB_Sanction_Letter_19092018.pdf"
OUT = Path(__file__).resolve().parent / "103_PNB_Sanction_Letter_19092018.md"
SOURCE_NAME = "103_PNB_Sanction_Letter_19092018.pdf"

# Four-page scanned sanction pack; pymupdf get_text() returns 0 chars per page.
# Transcriptions below from visual review of rendered pages (2x zoom).
SCANNED_TRANSCRIPTIONS: dict[int, str] = {
    1: """\
PNB Housing Finance Limited — Ghar Ki Baat
Branch Office: GURGAON
Date: 19-SEP-18

Applicant: Mr. Vijyant Agarwal, Mrs. Vineeta Agarwal
Ref No.: HOU/GUR/0918/578047
Version No.: 3

Address: FLAT NO 7, TYPE 5, SOUTH CAMPUS NETAJI SUBHAS, INSTITUTE OF TECH,
SECTOR 3, DWARKA, Delhi, DELHI-110075, India
Email: vijyantonly@yahoo.com
Contact No.: 9899308574

Dear Sir/Madam,

We are pleased to inform you that your application for a loan has been sanctioned
on the following terms and conditions:

No. | Field | Details
1. | Purpose of Loan | Housing Loan
2. | Loan Amount | INR 66,45,815.00 (INR Sixty-Six Lakhs Forty-Five Thousand Eight
   |             | Hundred Fifteen only)
   | Insurance Premium | INR 0.00 (INR Zero only)
   | Total Loan Amount | INR 66,45,815.00 (INR Sixty-Six Lakhs Forty-Five Thousand
   |                 | Eight Hundred Fifteen only)
3. | Repayment Term | 240 Months
4. | Rate of Interest Scheme | Floating
5. | PNBHFR * | 9.00% per annum
6. | Applicable Rate of Interest * | 9.00% + 0.10% = 9.10% per annum as on the date of
   |                             | execution of the loan agreement
7. | Equated Monthly Instalment ("EMI") * | INR 60,223.00 (INR Sixty Thousand Two
   |                                      | Hundred Twenty-Three only)
8. | Processing fees receivable: | INR 16,066.00 (INR Sixteen Thousand Sixty-Six only)
9. | Sanction letter validity | 90 days, from the date of this letter
10.| Security | 3051, 5, Tower 3, ATS Grandstand Phase 1, Sector 99A, Gurgaon,
   |          | Gurgaon, Haryana, India-122505

Special Conditions:
1. Offer subject to satisfactory legal and technical clearance of the property.
2. GPF to be reduced to Rs. 5k from 50K for both applicants; documentary proof required.
3. Form 26AS of co-applicant for AY 2017-18 & AY 2018-19 to be taken.
4. Dual signature affidavit of co-applicant for difference in signature between PAN Card
   and application form.
5. Sanction letter/repayment schedule of HL of Rs. 50.00 lacs availed from HDFC to be taken.
6. Loan to value (LTV) of property restricted to 75%.
7. Repayment from salary account of applicant or co-applicant.

Customer Service Manager: Pooja Jain (Landline: 0124 4055588)
Relationship Manager: GAURAV GUPTA-9910377657
Branch Address: SCO No. 391, Sector - 29, Urban Estate Near Iffco Chowk Metro Station,
Gurgaon-122001 Tel. 0124 4055588

For PNB Housing Finance Limited          Accepted all terms & conditions
[Authorised Signatory — signed/stamped]  [Borrower signatures — two]

Page 1 of 2

Regd. Office: 9th floor, Antriksh Bhavan, 22 Kasturba Gandhi Marg, New Delhi - 110 001
Phone: 011-23736857, Email: loans@pnbhfl.com, Website: www.pnbhfl.com
CIN: U65922DL1988PLC033856
""",
    2: """\
OTHER TERMS & CONDITIONS:

1. Loan Sanction: The sanction of the loan is subject to the Borrower(s) accepting all
   the terms and conditions mentioned herein and countersigning this letter. The Borrower(s)
   shall be deemed to have accepted all the terms and conditions upon countersigning.

2. Login/Processing Fee: Login fee and/or processing fee, as applicable, shall be payable
   by the Borrower(s) and shall be non-refundable under any circumstances.

3. Interest:
   (a) Fixed Rate: Where the loan is sanctioned at a fixed rate of interest, the rate shall
       remain fixed for the period specified in the loan agreement.
   (b) Floating Rate: Where the loan is sanctioned at a floating rate of interest, the rate
       shall be linked to PNBHFR (PNB Housing Floating Reference Rate) and shall be subject
       to revision from time to time.
   (c) Fixed & Floating Rate: Where applicable, the loan may carry a fixed rate for an
       initial period and thereafter convert to floating rate.
   (d) Option to convert between fixed and floating rate as per policy of PNBHFL.
   (e) Rate of interest shall be subject to revision as per the loan agreement.

4. Disbursement: The loan may be disbursed in lump sum or in installments as per the
   disbursement schedule agreed upon.

5. Repayment: The Borrower(s) shall repay the loan in Equated Monthly Installments (EMIs)
   and/or Pre-EMI Interest (PEMII) as applicable. Repayment shall be through Post Dated
   Cheques (PDCs) and/or ECS mandate as per PNBHFL requirements.

6. Prepayment Charges: Prepayment charges, if any, shall be as per guidelines of the
   National Housing Bank (NHB) and PNBHFL policy.

7. Security: The property offered as security shall be mortgaged in favour of PNBHFL as
   per the loan agreement.

8. Insurance: The Borrower(s) shall insure the property and life cover as required by PNBHFL.

9. NRI/PIO: In case of NRI/PIO borrowers, compliance with FEMA, 1999 and applicable RBI
   guidelines shall be ensured.

Miscellaneous

10. Change of address or status shall be intimated to PNBHFL in writing.
11. Loan amount shall not be used for any purpose other than stated in the application.
12. The sanction shall become null and void if:
    (a) Any material information is found to be incorrect or concealed;
    (b) There is deterioration in the financial position of the Borrower(s);
    (c) Any other event as per loan agreement occurs.
13. The Borrower(s) shall indemnify PNBHFL against any loss arising from breach of terms.
14. Corporate borrowers shall comply with the Companies Act, 1956/2013 as applicable.
15. PNBHFL reserves the right to revise interest rate on floating rate loans.
16. PNBHFL reserves the right not to disburse without assigning reasons.
17. PNBHFL's decision on all matters shall be final and binding.
18. This letter does not create any legal or equitable right until loan agreement is executed.
19. Stamp duty, registration and other charges shall be borne by the Borrower(s).
20. Validity: This sanction letter shall remain valid for 60 days from the date of issuance.

Version July, 2013

I/we have read/understood the terms and conditions mentioned herein and accord my/our
acceptance to the same.

Accepted all Terms & Conditions          (Borrower/s Name / Signature)
[Handwritten signatures — two]
""",
    3: """\
PNB Housing Finance Limited — Ghar Ki Baat

For PNB Housing Finance Limited          Borrower(s) (Name / Signature)
[Authorised Signatory — signed/stamped]  Vijyant Agarwal [signed]

Page 2 of 2

Regd. Office: 9th floor, Antriksh Bhavan, 22 Kasturba Gandhi Marg, New Delhi - 110 001
Phone: 011-23736857, Email: loans@pnbhfl.com, Website: www.pnbhfl.com
CIN: U65922DL1988PLC033856
""",
    4: """\
OTHER TERMS & CONDITIONS:

1. Loan Sanction: The sanction of the loan is subject to the Borrower(s) accepting all
   the terms and conditions mentioned herein and countersigning this letter. The Borrower(s)
   shall be deemed to have accepted all the terms and conditions upon countersigning.

2. Login/Processing Fee: Login fee and/or processing fee, as applicable, shall be payable
   by the Borrower(s) and shall be non-refundable under any circumstances.

3. Interest:
   (a) Fixed Rate: Where the loan is sanctioned at a fixed rate of interest, the rate shall
       remain fixed for the period specified in the loan agreement.
   (b) Floating Rate: Where the loan is sanctioned at a floating rate of interest, the rate
       shall be linked to PNBHFR (PNB Housing Floating Reference Rate) and shall be subject
       to revision from time to time.
   (c) Fixed & Floating Rate: Where applicable, the loan may carry a fixed rate for an
       initial period and thereafter convert to floating rate.
   (d) Option to convert between fixed and floating rate as per policy of PNBHFL.
   (e) Rate of interest shall be subject to revision as per the loan agreement.

4. Disbursement: The loan may be disbursed in lump sum or in installments as per the
   disbursement schedule agreed upon.

5. Repayment: The Borrower(s) shall repay the loan in Equated Monthly Installments (EMIs)
   and/or Pre-EMI Interest (PMII) as applicable. Repayment shall be through Post Dated
   Cheques (PDCs) and/or ECS mandate as per PNBHFL requirements.

6. Prepayment Charges: Prepayment charges, if any, shall be as per guidelines of the
   National Housing Bank (NHB) and PNBHFL policy.

7. Security: The property offered as security shall be mortgaged in favour of PNBHFL as
   per the loan agreement.

8. Insurance: The Borrower(s) shall insure the property and life cover as required by PNBHFL.

9. NRI/PIO: In case of NRI/PIO borrowers, compliance with FEMA, 1999 and applicable RBI
   guidelines shall be ensured.

Miscellaneous

10. Change of address or status shall be intimated to PNBHFL in writing.
11. Loan amount shall not be used for any purpose other than stated in the application.
12. The sanction shall become null and void if:
    (a) Any material information is found to be incorrect or concealed;
    (b) There is deterioration in the financial position of the Borrower(s);
    (c) Any other event as per loan agreement occurs.
13. The Borrower(s) shall indemnify PNBHFL against any loss arising from breach of terms.
14. Corporate borrowers shall comply with the Companies Act, 1956/2013 as applicable.
15. PNBHFL reserves the right to revise interest rate on floating rate loans.
16. PNBHFL reserves the right not to disburse without assigning reasons.
17. PNBHFL's decision on all matters shall be final and binding.
18. This letter does not create any legal or equitable right until loan agreement is executed.
19. Stamp duty, registration and other charges shall be borne by the Borrower(s).
20. Validity: This sanction letter shall remain valid for 60 days from the date of issuance.

Version July, 2013

I/we have read/understood the terms and conditions mentioned herein and accord my/our
acceptance to the same.

Accepted all Terms & Conditions          (Borrower/s Name / Signature)
Vijyant Agarwal [signed]
""",
}


def has_subvention(text: str) -> bool:
    return "subvention" in text.lower()


def main() -> None:
    doc = fitz.open(PDF)
    page_count = doc.page_count
    text_pages = 0
    image_only_pages = 0
    all_text_parts: list[str] = []

    parts = [
        "# PNB Housing — Sanction letter (19-Sep-2018)\n\n",
        f"Converted from: `{SOURCE_NAME}`\n\n",
        "**Note:** Loan sanction letter dated **19-SEP-2018** from PNB Housing Finance "
        "Limited, Gurgaon branch, for housing loan to Mr. Vijyant Agarwal & "
        "Mrs. Vineeta Agarwal (Unit 3051, ATS Grandstand Phase 1).\n\n",
        f"**Pages:** {page_count}\n\n---\n",
    ]

    for i, page in enumerate(doc, start=1):
        text = page.get_text().strip()
        has_images = bool(page.get_images())
        parts.append(f"\n## Page {i}\n\n")
        all_text_parts.append(text)

        if text:
            text_pages += 1
            parts.append("```text\n" + text + "\n```\n")
        elif has_images and i in SCANNED_TRANSCRIPTIONS:
            image_only_pages += 1
            transcription = SCANNED_TRANSCRIPTIONS[i].strip()
            all_text_parts.append(transcription)
            parts.append(
                "*(Scanned/image-only page — no extractable text layer in PDF. "
                "Transcription below from visual review of source.)*\n\n"
            )
            parts.append("```text\n" + transcription + "\n```\n")
        else:
            parts.append("*(No extractable text.)*\n")

    doc.close()

    combined = "\n".join(all_text_parts)
    subvention_found = has_subvention(combined)

    parts.append(
        "\n---\n\n## Key details (extracted)\n\n"
        "| Field | Value |\n"
        "|-------|--------|\n"
        "| **Sanction date** | 19-SEP-2018 |\n"
        "| **Lender** | PNB Housing Finance Limited (PNBHFL) |\n"
        "| **Branch** | Gurgaon (SCO No. 391, Sector 29) |\n"
        "| **Ref No.** | HOU/GUR/0918/578047 |\n"
        "| **Borrowers** | Mr. Vijyant Agarwal & Mrs. Vineeta Agarwal |\n"
        "| **Purpose** | Housing Loan |\n"
        "| **Sanction amount** | Rs. 66,45,815.00 |\n"
        "| **Repayment term** | 240 months |\n"
        "| **Interest scheme** | Floating (PNBHFR 9.00% + 0.10% = 9.10% p.a.) |\n"
        "| **EMI (sanction letter)** | Rs. 60,223.00 per month |\n"
        "| **Processing fees** | Rs. 16,066.00 |\n"
        "| **Sanction validity** | 90 days (main letter); T&C page notes 60 days |\n"
        "| **Security / property** | Unit 3051, Tower 3, ATS Grandstand Phase 1, Sector 99A, Gurgaon |\n"
        "| **LTV restriction** | 75% |\n"
        "| **Pre-EMI / PEMII** | Mentioned in T&C clause 5 (repayment terms) |\n"
        f"| **Subvention (word in document)** | {'Yes' if subvention_found else 'No'} |\n"
        "| **Relationship Manager** | Gaurav Gupta (9910377657) |\n"
        "| **Customer Service Manager** | Pooja Jain |\n"
        "\n## Extraction note\n\n"
        f"- **Source PDF:** `attachments/{SOURCE_NAME}`\n"
        f"- **Total pages:** {page_count}\n"
        f"- **Pages with extractable text:** {text_pages}\n"
        f"- **Scanned/image-only pages:** {image_only_pages} "
        "(embedded JPEG; pymupdf get_text() returned 0 chars per page)\n"
        "- **Document type:** Scanned/image-only (not text-layer PDF)\n"
        f"- **Subvention mentioned:** {'Yes' if subvention_found else 'No'} "
        "(word 'subvention' not found; Pre-EMI/PEMII referenced in T&C only)\n"
        "- **Transcription:** All pages transcribed manually from rendered page images "
        "because the source PDF has no extractable text layer.\n"
        "- **Duplicate content:** Pages 2 and 4 both contain the standard "
        "'Other Terms & Conditions' schedule (appears to be duplicate scan).\n"
    )

    OUT.write_text("".join(parts), encoding="utf-8")
    print(
        f"Wrote {OUT} ({OUT.stat().st_size} bytes, {page_count} pages, "
        f"subvention={'yes' if subvention_found else 'no'}, "
        f"text_pages={text_pages}, scanned_pages={image_only_pages})"
    )


if __name__ == "__main__":
    main()
