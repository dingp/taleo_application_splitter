# Application Splitter

This directory contains three Python scripts for working with candidate application PDFs.

## Files

- `split_applications.py`
  Splits a batch PDF into one PDF per applicant.
- `extract_applicants_to_xlsx.py`
  Extracts applicant name, applicant ID, and email from individual applicant PDFs into an `.xlsx` file.
- `update_email_columns.py`
  Uses `applicants.xlsx` to fill the email column in `Candidate Summary.xlsx` and `Candidates to phone screen.xlsx`.

## Requirements

These scripts use:

- `python3`
- `pdftotext`
- `pdfinfo`
- `pdfseparate`
- `pdfunite`

On this machine, those PDF tools are available from Poppler.

## 1. Split batch PDFs into one PDF per applicant

Run on one batch PDF:

```bash
cd /Users/dingpf/Desktop/dseg_hire/application_splitter
python3 split_applications.py "../batch_1.pdf"
```

That creates a directory next to the source PDF named:

```text
../batch_1_split
```

Run on multiple batch PDFs and send outputs to one base directory:

```bash
cd /Users/dingpf/Desktop/dseg_hire/application_splitter
python3 split_applications.py \
  "../batch_1.pdf" \
  "../batch_2.pdf" \
  "../batch_3.pdf" \
  -o ../all_split \
  --validate
```

Notes:

- Output filenames are based on applicant name and applicant ID, for example:
  `Applicant_Name_ID.pdf`
- `--validate` rescans created PDFs and prints a warning if any output still appears to contain more than one applicant.

## 2. Extract applicant name and email into `applicants.xlsx`

Extract from all individual applicant PDFs under a split directory:

```bash
cd /Users/dingpf/Desktop/dseg_hire/application_splitter
python3 extract_applicants_to_xlsx.py ../all_split_corrected/**/*.pdf -o ../applicants.xlsx
```

If your shell does not expand `**`, use `find`:

```bash
cd /Users/dingpf/Desktop/dseg_hire/application_splitter
python3 extract_applicants_to_xlsx.py $(find ../all_split_corrected -name '*.pdf' | sort) -o ../applicants.xlsx
```

Extract from only one batch’s split output:

```bash
cd /Users/dingpf/Desktop/dseg_hire/application_splitter
python3 extract_applicants_to_xlsx.py \
  ../all_split/batch_1/*.pdf \
  -o ../applicants_batch_1.xlsx
```

The output workbook contains:

- `Source File`
- `Applicant Name`
- `Applicant ID`
- `Email Address`

## 3. Update the review spreadsheets with email addresses

This script expects these files to exist in the parent directory:

- `../applicants.xlsx`
- `../Candidate Summary.xlsx`
- `../Candidates to phone screen.xlsx`

Run:

```bash
cd /Users/dingpf/Desktop/dseg_hire/application_splitter
python3 update_email_columns.py
```

This updates:

- `../Candidate Summary.xlsx`
- `../Candidates to phone screen.xlsx`

using the emails from `../applicants.xlsx`.

## Typical workflow

```bash
cd /Users/dingpf/Desktop/dseg_hire/application_splitter

python3 split_applications.py \
  "../batch_1.pdf" \
  "../batch_2.pdf" \
  "../batch_3.pdf" \
  -o ../all_split \
  --validate

python3 extract_applicants_to_xlsx.py $(find ../all_split -name '*.pdf' | sort) -o ../applicants.xlsx

python3 update_email_columns.py
```

## Output locations

- Split applicant PDFs: `../all_split/`
- Applicant spreadsheet: `../applicants.xlsx`
- Updated review workbooks:
  `../Candidate Summary.xlsx`
  and
  `../Candidates to phone screen.xlsx`
