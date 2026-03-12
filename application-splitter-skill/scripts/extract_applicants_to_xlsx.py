#!/usr/bin/env python3
"""Extract applicant names and emails from individual PDFs into an XLSX file."""

from __future__ import annotations

import argparse
import datetime as dt
import re
import shutil
import subprocess
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape


NAME_PATTERN = re.compile(
    r"(?P<name>[^\n]+?)\s*\((?P<applicant_id>\d+)\)\s*applied\s+for\s+job:",
    re.IGNORECASE | re.DOTALL,
)
EMAIL_PATTERN = re.compile(r"\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b", re.IGNORECASE)


def require_tool(name: str) -> str:
    path = shutil.which(name)
    if not path:
        raise SystemExit(f"Required tool not found in PATH: {name}")
    return path


def run_command(args: list[str]) -> str:
    try:
        result = subprocess.run(args, check=True, capture_output=True, text=True)
    except subprocess.CalledProcessError as exc:
        stderr = exc.stderr.strip() if exc.stderr else ""
        raise SystemExit(f"Command failed: {' '.join(args)}\n{stderr}") from exc
    return result.stdout


def pdf_text(pdf_path: Path) -> str:
    return run_command([require_tool("pdftotext"), str(pdf_path), "-"])


def normalize_space(value: str) -> str:
    return " ".join(value.split())


def extract_record(pdf_path: Path) -> dict[str, str]:
    text = pdf_text(pdf_path)
    name_match = NAME_PATTERN.search(text)
    email_match = EMAIL_PATTERN.search(text)

    return {
        "source_file": pdf_path.name,
        "name": normalize_space(name_match.group("name")) if name_match else "",
        "applicant_id": name_match.group("applicant_id") if name_match else "",
        "email": email_match.group(0) if email_match else "",
    }


def excel_column_name(index: int) -> str:
    letters = []
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        letters.append(chr(65 + remainder))
    return "".join(reversed(letters))


def shared_strings_xml(strings: list[str]) -> str:
    items = "".join(f"<si><t>{escape(value)}</t></si>" for value in strings)
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        f'count="{len(strings)}" uniqueCount="{len(strings)}">{items}</sst>'
    )


def worksheet_xml(rows: list[list[int]]) -> str:
    row_xml = []
    for row_number, row in enumerate(rows, start=1):
        cells = []
        for column_number, string_index in enumerate(row, start=1):
            cell_ref = f"{excel_column_name(column_number)}{row_number}"
            cells.append(f'<c r="{cell_ref}" t="s"><v>{string_index}</v></c>')
        row_xml.append(f'<row r="{row_number}">{"".join(cells)}</row>')
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<sheetViews><sheetView workbookViewId="0"/></sheetViews>'
        '<sheetFormatPr defaultRowHeight="15"/>'
        '<cols>'
        '<col min="1" max="1" width="32" customWidth="1"/>'
        '<col min="2" max="2" width="28" customWidth="1"/>'
        '<col min="3" max="3" width="16" customWidth="1"/>'
        '<col min="4" max="4" width="36" customWidth="1"/>'
        '</cols>'
        f'<sheetData>{"".join(row_xml)}</sheetData>'
        "</worksheet>"
    )


def workbook_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheets><sheet name="Applicants" sheetId="1" r:id="rId1"/></sheets>'
        "</workbook>"
    )


def content_types_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '<Override PartName="/xl/sharedStrings.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
        '<Override PartName="/docProps/core.xml" '
        'ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
        '<Override PartName="/docProps/app.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
        "</Types>"
    )


def root_rels_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="xl/workbook.xml"/>'
        '<Relationship Id="rId2" '
        'Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" '
        'Target="docProps/core.xml"/>'
        '<Relationship Id="rId3" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" '
        'Target="docProps/app.xml"/>'
        "</Relationships>"
    )


def workbook_rels_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
        'Target="worksheets/sheet1.xml"/>'
        '<Relationship Id="rId2" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" '
        'Target="sharedStrings.xml"/>'
        "</Relationships>"
    )


def app_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
        'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
        '<Application>Codex</Application>'
        "</Properties>"
    )


def core_xml() -> str:
    created = dt.datetime.now(dt.timezone.utc).replace(microsecond=0).isoformat()
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/" '
        'xmlns:dcterms="http://purl.org/dc/terms/" '
        'xmlns:dcmitype="http://purl.org/dc/dcmitype/" '
        'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        '<dc:creator>Codex</dc:creator>'
        '<cp:lastModifiedBy>Codex</cp:lastModifiedBy>'
        f'<dcterms:created xsi:type="dcterms:W3CDTF">{created}</dcterms:created>'
        f'<dcterms:modified xsi:type="dcterms:W3CDTF">{created}</dcterms:modified>'
        "</cp:coreProperties>"
    )


def write_xlsx(records: list[dict[str, str]], output_path: Path) -> None:
    rows = [["Source File", "Applicant Name", "Applicant ID", "Email Address"]]
    for record in records:
        rows.append(
            [
                record["source_file"],
                record["name"],
                record["applicant_id"],
                record["email"],
            ]
        )

    shared_strings: list[str] = []
    index_map: dict[str, int] = {}
    worksheet_rows: list[list[int]] = []
    for row in rows:
        worksheet_row = []
        for value in row:
            if value not in index_map:
                index_map[value] = len(shared_strings)
                shared_strings.append(value)
            worksheet_row.append(index_map[value])
        worksheet_rows.append(worksheet_row)

    with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", content_types_xml())
        archive.writestr("_rels/.rels", root_rels_xml())
        archive.writestr("docProps/app.xml", app_xml())
        archive.writestr("docProps/core.xml", core_xml())
        archive.writestr("xl/workbook.xml", workbook_xml())
        archive.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml())
        archive.writestr("xl/sharedStrings.xml", shared_strings_xml(shared_strings))
        archive.writestr("xl/worksheets/sheet1.xml", worksheet_xml(worksheet_rows))


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Extract applicant names and emails from individual PDFs into an XLSX file."
    )
    parser.add_argument("inputs", nargs="+", type=Path, help="Individual applicant PDF files.")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=Path("applicants.xlsx"),
        help="Output XLSX path. Defaults to ./applicants.xlsx",
    )
    args = parser.parse_args()

    records = []
    for pdf_path in args.inputs:
        if pdf_path.suffix.lower() != ".pdf":
            raise SystemExit(f"Input is not a PDF: {pdf_path}")
        if not pdf_path.exists():
            raise SystemExit(f"Input does not exist: {pdf_path}")
        records.append(extract_record(pdf_path))

    write_xlsx(records, args.output)
    print(f"Wrote {len(records)} rows to {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
