---
name: application-splitter
description: Split batched applicant PDF packets into one PDF per applicant, extract applicant name and email from individual PDFs into an XLSX workbook, and update existing candidate-review workbooks with email columns. Use this when working with recruiter export PDFs and spreadsheets for job applications.
---

# Application Splitter

Use this skill when the task involves recruiter-exported PDF packets where each applicant starts on a recognizable summary page, plus downstream extraction of applicant metadata into spreadsheets.

## Included scripts

- `scripts/split_applications.py`
  Splits one or more batched application PDFs into one PDF per applicant.
- `scripts/extract_applicants_to_xlsx.py`
  Extracts applicant name, applicant ID, and email from individual applicant PDFs into an `.xlsx`.
- `scripts/update_email_columns.py`
  Uses `applicants.xlsx` to fill email columns in candidate review workbooks.

## Runtime requirements

- `python3` with Python 3.9+ recommended
- Poppler command-line tools:
  - `pdftotext`
  - `pdfinfo`
  - `pdfseparate`
  - `pdfunite`

These scripts do not require third-party Python packages. They use the Python standard library only.

Tool mapping:

- `scripts/split_applications.py` needs `pdftotext`, `pdfinfo`, `pdfseparate`, and `pdfunite`
- `scripts/extract_applicants_to_xlsx.py` needs `pdftotext`
- `scripts/update_email_columns.py` needs only Python standard library modules

## Workflow

1. Split the batch PDFs.
2. Build `applicants.xlsx` from the individual applicant PDFs.
3. Update the review workbooks from `applicants.xlsx`.

## Commands

Split one or more batch PDFs:

```bash
python3 scripts/split_applications.py "../batch_1.pdf" -o ../all_split --validate
python3 scripts/split_applications.py "../batch_1.pdf" "../batch_2.pdf" "../batch_3.pdf" -o ../all_split --validate
```

Generate `applicants.xlsx` from split PDFs:

```bash
python3 scripts/extract_applicants_to_xlsx.py $(find ../all_split -name '*.pdf' | sort) -o ../applicants.xlsx
```

Update review spreadsheets:

```bash
python3 scripts/update_email_columns.py
```

## Notes

- The splitter detects applicant boundaries from the applicant summary page text.
- `--validate` is useful after a split to flag any output PDF that still appears to contain more than one applicant.
- The workbook updater expects `applicants.xlsx`, `Candidate Summary.xlsx`, and `Candidates to phone screen.xlsx` in the working directory’s parent or another path you choose to manage externally.

## Reference

For concrete usage examples, file layout, and end-to-end command sequences, see [README.md](README.md).
