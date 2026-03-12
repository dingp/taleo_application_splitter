#!/usr/bin/env python3
"""Populate email columns in the candidate workbooks from applicants.xlsx."""

from __future__ import annotations

import copy
import re
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS = {"a": MAIN_NS}
ET.register_namespace("", MAIN_NS)
ET.register_namespace("r", REL_NS)

CELL_REF_RE = re.compile(r"([A-Z]+)(\d+)")
NAME_ID_RE = re.compile(r"^(.*?)\s*\((\d+)\)?\s*$")


def col_to_index(col: str) -> int:
    value = 0
    for char in col:
        value = value * 26 + (ord(char) - 64)
    return value


def index_to_col(index: int) -> str:
    letters: list[str] = []
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        letters.append(chr(65 + remainder))
    return "".join(reversed(letters))


def split_ref(ref: str) -> tuple[str, int]:
    match = CELL_REF_RE.fullmatch(ref)
    if not match:
        raise ValueError(f"Invalid cell reference: {ref}")
    return match.group(1), int(match.group(2))


def shift_ref(ref: str, start_col_index: int, delta: int) -> str:
    col, row = split_ref(ref)
    col_index = col_to_index(col)
    if col_index >= start_col_index:
        col_index += delta
    return f"{index_to_col(col_index)}{row}"


def normalize_text(value: str) -> str:
    return " ".join((value or "").split())


def normalize_name(value: str) -> tuple[str, str]:
    cleaned = normalize_text(value)
    match = NAME_ID_RE.match(cleaned)
    if match:
        return normalize_text(match.group(1)).casefold(), match.group(2)
    return cleaned.casefold(), ""


def parse_shared_strings(root: ET.Element) -> list[str]:
    values: list[str] = []
    for item in root.findall("a:si", NS):
        values.append("".join(node.text or "" for node in item.iterfind(".//a:t", NS)))
    return values


def cell_value(cell: ET.Element, shared_strings: list[str]) -> str:
    cell_type = cell.attrib.get("t")
    if cell_type == "s":
        value = cell.findtext(f"{{{MAIN_NS}}}v", default="")
        return shared_strings[int(value)] if value else ""
    if cell_type == "inlineStr":
        return "".join(node.text or "" for node in cell.iterfind(".//a:t", NS))
    return cell.findtext(f"{{{MAIN_NS}}}v", default="")


def load_email_lookup(
    applicants_path: Path,
) -> tuple[dict[tuple[str, str], str], dict[str, str]]:
    with zipfile.ZipFile(applicants_path) as archive:
        shared_strings = parse_shared_strings(
            ET.fromstring(archive.read("xl/sharedStrings.xml"))
        )
        worksheet = ET.fromstring(archive.read("xl/worksheets/sheet1.xml"))

    by_name_and_id: dict[tuple[str, str], str] = {}
    by_name: dict[str, str] = {}
    for row in worksheet.find(f"{{{MAIN_NS}}}sheetData").findall(f"{{{MAIN_NS}}}row"):
        values = {}
        for cell in row.findall(f"{{{MAIN_NS}}}c"):
            col, _ = split_ref(cell.attrib["r"])
            values[col] = cell_value(cell, shared_strings)
        if values.get("B") == "Applicant Name":
            continue
        name_key, _ = normalize_name(values.get("B", ""))
        applicant_id = normalize_text(values.get("C", ""))
        email = values.get("D", "")
        if name_key:
            by_name[name_key] = email
            if applicant_id:
                by_name_and_id[(name_key, applicant_id)] = email
    return by_name_and_id, by_name


def make_inline_string_cell(ref: str, style: str | None, text: str) -> ET.Element:
    cell = ET.Element(f"{{{MAIN_NS}}}c", {"r": ref, "t": "inlineStr"})
    if style is not None:
        cell.set("s", style)
    inline = ET.SubElement(cell, f"{{{MAIN_NS}}}is")
    text_node = ET.SubElement(inline, f"{{{MAIN_NS}}}t")
    text_node.text = text
    return cell


def row_cell_map(row: ET.Element) -> dict[str, ET.Element]:
    return {split_ref(cell.attrib["r"])[0]: cell for cell in row.findall(f"{{{MAIN_NS}}}c")}


def style_for_row(row: ET.Element, preferred_col: str, fallback_col: str) -> str | None:
    cells = row_cell_map(row)
    for col in (preferred_col, fallback_col):
        cell = cells.get(col)
        if cell is not None and "s" in cell.attrib:
            return cell.attrib["s"]
    return None


def email_for_name(
    raw_name: str,
    by_name_and_id: dict[tuple[str, str], str],
    by_name: dict[str, str],
) -> str:
    name_key, applicant_id = normalize_name(raw_name)
    if not name_key:
        return ""
    if applicant_id and (name_key, applicant_id) in by_name_and_id:
        return by_name_and_id[(name_key, applicant_id)]
    return by_name.get(name_key, "")


def update_cols_for_insert(worksheet: ET.Element) -> None:
    cols = worksheet.find(f"{{{MAIN_NS}}}cols")
    if cols is None:
        return
    new_cols: list[ET.Element] = []
    inserted = False
    for col in cols.findall(f"{{{MAIN_NS}}}col"):
        col_copy = copy.deepcopy(col)
        min_value = int(col_copy.attrib["min"])
        max_value = int(col_copy.attrib["max"])
        if min_value >= 3:
            col_copy.set("min", str(min_value + 1))
            col_copy.set("max", str(max_value + 1))
        if not inserted and min_value >= 3:
            new_cols.append(
                ET.Element(
                    f"{{{MAIN_NS}}}col",
                    {"min": "3", "max": "3", "width": "15.25", "customWidth": "1"},
                )
            )
            inserted = True
        new_cols.append(col_copy)
    if not inserted:
        new_cols.append(
            ET.Element(
                f"{{{MAIN_NS}}}col",
                {"min": "3", "max": "3", "width": "15.25", "customWidth": "1"},
            )
        )
    cols[:] = new_cols


def shift_hyperlinks(worksheet: ET.Element) -> None:
    hyperlinks = worksheet.find(f"{{{MAIN_NS}}}hyperlinks")
    if hyperlinks is None:
        return
    for hyperlink in hyperlinks.findall(f"{{{MAIN_NS}}}hyperlink"):
        hyperlink.set("ref", shift_ref(hyperlink.attrib["ref"], 3, 1))


def resolve_row_values(
    worksheet: ET.Element, shared_strings: list[str]
) -> dict[int, dict[str, str]]:
    resolved: dict[int, dict[str, str]] = {}
    sheet_data = worksheet.find(f"{{{MAIN_NS}}}sheetData")
    if sheet_data is None:
        return resolved
    for row in sheet_data.findall(f"{{{MAIN_NS}}}row"):
        row_number = int(row.attrib["r"])
        resolved[row_number] = {}
        for cell in row.findall(f"{{{MAIN_NS}}}c"):
            col, _ = split_ref(cell.attrib["r"])
            resolved[row_number][col] = cell_value(cell, shared_strings)
    return resolved


def has_email_column(worksheet: ET.Element, shared_strings: list[str]) -> bool:
    resolved = resolve_row_values(worksheet, shared_strings)
    return normalize_text(resolved.get(1, {}).get("C", "")).casefold() == "email"


def insert_email_column_with_values(
    worksheet: ET.Element,
    shared_strings: list[str],
    email_lookup: tuple[dict[tuple[str, str], str], dict[str, str]],
) -> None:
    by_name_and_id, by_name = email_lookup
    resolved = resolve_row_values(worksheet, shared_strings)
    sheet_data = worksheet.find(f"{{{MAIN_NS}}}sheetData")
    if sheet_data is None:
        return

    for row in sheet_data.findall(f"{{{MAIN_NS}}}row"):
        row_number = int(row.attrib["r"])
        cells = row.findall(f"{{{MAIN_NS}}}c")
        for cell in reversed(cells):
            col, _ = split_ref(cell.attrib["r"])
            if col_to_index(col) >= 3:
                cell.attrib["r"] = shift_ref(cell.attrib["r"], 3, 1)

        style = style_for_row(row, "C", "B")
        if row_number == 1:
            email_value = "Email"
        else:
            email_value = email_for_name(
                resolved.get(row_number, {}).get("B", ""), by_name_and_id, by_name
            )
        if email_value:
            insert_at = 0
            children = row.findall(f"{{{MAIN_NS}}}c")
            while insert_at < len(children):
                col, _ = split_ref(children[insert_at].attrib["r"])
                if col_to_index(col) > 3:
                    break
                insert_at += 1
            row.insert(insert_at, make_inline_string_cell(f"C{row_number}", style, email_value))

    update_cols_for_insert(worksheet)
    shift_hyperlinks(worksheet)


def populate_existing_email_column(
    worksheet: ET.Element,
    shared_strings: list[str],
    email_lookup: tuple[dict[tuple[str, str], str], dict[str, str]],
) -> None:
    by_name_and_id, by_name = email_lookup
    resolved = resolve_row_values(worksheet, shared_strings)
    sheet_data = worksheet.find(f"{{{MAIN_NS}}}sheetData")
    if sheet_data is None:
        return

    for row in sheet_data.findall(f"{{{MAIN_NS}}}row"):
        row_number = int(row.attrib["r"])
        cells = row_cell_map(row)
        if row_number == 1:
            email_value = "Email"
            style = style_for_row(row, "C", "B")
        else:
            email_value = email_for_name(
                resolved.get(row_number, {}).get("B", ""), by_name_and_id, by_name
            )
            style = style_for_row(row, "C", "B")
        target = cells.get("C")
        if target is None:
            if not email_value:
                continue
            insert_at = 0
            children = row.findall(f"{{{MAIN_NS}}}c")
            while insert_at < len(children):
                col, _ = split_ref(children[insert_at].attrib["r"])
                if col_to_index(col) > 3:
                    break
                insert_at += 1
            row.insert(insert_at, make_inline_string_cell(f"C{row_number}", style, email_value))
            continue

        target.attrib["r"] = f"C{row_number}"
        target.attrib["t"] = "inlineStr"
        for child in list(target):
            target.remove(child)
        inline = ET.SubElement(target, f"{{{MAIN_NS}}}is")
        text_node = ET.SubElement(inline, f"{{{MAIN_NS}}}t")
        text_node.text = email_value


def rewrite_workbook(
    workbook_path: Path,
    updater,
    email_lookup: tuple[dict[tuple[str, str], str], dict[str, str]],
) -> None:
    with zipfile.ZipFile(workbook_path) as source:
        worksheet = ET.fromstring(source.read("xl/worksheets/sheet1.xml"))
        shared_strings = parse_shared_strings(ET.fromstring(source.read("xl/sharedStrings.xml")))
        updater(worksheet, shared_strings, email_lookup)
        worksheet_bytes = ET.tostring(worksheet, encoding="utf-8", xml_declaration=True)

        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx", dir=workbook_path.parent) as temp:
            temp_path = Path(temp.name)

        try:
            with zipfile.ZipFile(temp_path, "w", compression=zipfile.ZIP_DEFLATED) as target:
                for item in source.infolist():
                    if item.filename == "xl/worksheets/sheet1.xml":
                        target.writestr(item, worksheet_bytes)
                    else:
                        target.writestr(item, source.read(item.filename))
            temp_path.replace(workbook_path)
        finally:
            if temp_path.exists():
                temp_path.unlink()


def update_candidate_summary(
    worksheet: ET.Element,
    shared_strings: list[str],
    email_lookup: tuple[dict[tuple[str, str], str], dict[str, str]],
) -> None:
    if has_email_column(worksheet, shared_strings):
        populate_existing_email_column(worksheet, shared_strings, email_lookup)
    else:
        insert_email_column_with_values(worksheet, shared_strings, email_lookup)


def main() -> int:
    email_lookup = load_email_lookup(Path("applicants.xlsx"))
    rewrite_workbook(
        Path("Candidate Summary.xlsx"),
        update_candidate_summary,
        email_lookup,
    )
    rewrite_workbook(
        Path("Candidates to phone screen.xlsx"),
        populate_existing_email_column,
        email_lookup,
    )
    print("Updated email columns in Candidate Summary.xlsx and Candidates to phone screen.xlsx")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
