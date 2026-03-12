#!/usr/bin/env python3
"""Split batched applicant PDFs into one PDF per applicant."""

from __future__ import annotations

import argparse
import re
import shutil
import subprocess
import tempfile
import unicodedata
from dataclasses import dataclass
from pathlib import Path


START_PATTERN = re.compile(
    r"(?P<name>[^\n]+?)\s*\((?P<applicant_id>\d+)\)\s*applied\s+for\s+job:",
    re.IGNORECASE | re.DOTALL,
)


@dataclass(frozen=True)
class ApplicantSection:
    name: str
    applicant_id: str
    start_page: int
    end_page: int


def require_tool(name: str) -> str:
    path = shutil.which(name)
    if not path:
        raise SystemExit(f"Required tool not found in PATH: {name}")
    return path


def run_command(args: list[str], *, text: bool = False) -> str:
    try:
        result = subprocess.run(
            args,
            check=True,
            capture_output=True,
            text=text,
        )
    except subprocess.CalledProcessError as exc:
        stderr = exc.stderr.strip() if exc.stderr else ""
        raise SystemExit(f"Command failed: {' '.join(args)}\n{stderr}") from exc
    return result.stdout if text else ""


def pdf_page_count(pdf_path: Path) -> int:
    output = run_command([require_tool("pdfinfo"), str(pdf_path)], text=True)
    for line in output.splitlines():
        if line.startswith("Pages:"):
            return int(line.split(":", 1)[1].strip())
    raise SystemExit(f"Could not determine page count for {pdf_path}")


def extract_pages_text(pdf_path: Path) -> list[str]:
    output = run_command([require_tool("pdftotext"), str(pdf_path), "-"], text=True)
    page_count = pdf_page_count(pdf_path)
    pages = output.split("\f")
    if len(pages) < page_count:
        raise SystemExit(
            f"Expected at least {page_count} text pages from {pdf_path}, got {len(pages)}"
        )
    return pages[:page_count]


def normalize_space(value: str) -> str:
    return " ".join(value.split())


def safe_filename(name: str) -> str:
    normalized = unicodedata.normalize("NFKD", name)
    ascii_name = normalized.encode("ascii", "ignore").decode("ascii")
    ascii_name = ascii_name.replace(",", "")
    ascii_name = re.sub(r"[^A-Za-z0-9._ -]+", "", ascii_name)
    ascii_name = re.sub(r"\s+", "_", ascii_name.strip())
    return ascii_name or "applicant"


def find_sections(page_texts: list[str]) -> list[ApplicantSection]:
    starts: list[tuple[int, str, str]] = []
    for page_number, page_text in enumerate(page_texts, start=1):
        match = START_PATTERN.search(page_text)
        if not match:
            continue
        starts.append(
            (
                page_number,
                normalize_space(match.group("name")),
                match.group("applicant_id"),
            )
        )

    if not starts:
        raise SystemExit("No applicant start pages were detected.")

    sections: list[ApplicantSection] = []
    last_page = len(page_texts)
    for index, (start_page, name, applicant_id) in enumerate(starts):
        next_start = starts[index + 1][0] if index + 1 < len(starts) else last_page + 1
        sections.append(
            ApplicantSection(
                name=name,
                applicant_id=applicant_id,
                start_page=start_page,
                end_page=next_start - 1,
            )
        )
    return sections


def ensure_unique_path(path: Path) -> Path:
    if not path.exists():
        return path
    stem = path.stem
    suffix = path.suffix
    counter = 2
    while True:
        candidate = path.with_name(f"{stem}_{counter}{suffix}")
        if not candidate.exists():
            return candidate
        counter += 1


def split_pdf(pdf_path: Path, output_dir: Path) -> list[Path]:
    page_texts = extract_pages_text(pdf_path)
    sections = find_sections(page_texts)
    output_dir.mkdir(parents=True, exist_ok=True)

    created_files: list[Path] = []
    with tempfile.TemporaryDirectory(prefix="split_pages_") as temp_dir:
        temp_dir_path = Path(temp_dir)
        page_pattern = temp_dir_path / "page-%04d.pdf"
        run_command(
            [
                require_tool("pdfseparate"),
                str(pdf_path),
                str(page_pattern),
            ]
        )

        for section in sections:
            base_name = f"{safe_filename(section.name)}_{section.applicant_id}.pdf"
            destination = ensure_unique_path(output_dir / base_name)
            input_pages = [
                str(temp_dir_path / f"page-{page_number:04d}.pdf")
                for page_number in range(section.start_page, section.end_page + 1)
            ]
            run_command([require_tool("pdfunite"), *input_pages, str(destination)])
            created_files.append(destination)

    return created_files


def validate_outputs(output_files: list[Path]) -> list[str]:
    problems: list[str] = []
    for pdf_path in output_files:
        sections = find_sections(extract_pages_text(pdf_path))
        if len(sections) > 1:
            names = ", ".join(f"{section.name} ({section.applicant_id})" for section in sections)
            problems.append(f"{pdf_path} still contains multiple applicants: {names}")
    return problems


def build_output_dir(
    pdf_path: Path, base_output_dir: Path | None, input_count: int
) -> Path:
    if base_output_dir is None:
        return pdf_path.parent / f"{pdf_path.stem}_split"
    if input_count == 1:
        return base_output_dir
    return base_output_dir / pdf_path.stem


parser = argparse.ArgumentParser(
    description="Split batched application PDFs into one PDF per applicant."
)
parser.add_argument("inputs", nargs="+", type=Path, help="Input batched PDF files.")
parser.add_argument(
    "-o",
    "--output-dir",
    type=Path,
    help="Directory for output PDFs. With multiple inputs, subdirectories are created per source PDF.",
)
parser.add_argument(
    "--validate",
    action="store_true",
    help="Re-scan created output PDFs and warn if any still contain multiple applicants.",
)


def main(args: argparse.Namespace) -> int:
    for pdf_path in args.inputs:
        if pdf_path.suffix.lower() != ".pdf":
            raise SystemExit(f"Input is not a PDF: {pdf_path}")
        if not pdf_path.exists():
            raise SystemExit(f"Input does not exist: {pdf_path}")

    for pdf_path in args.inputs:
        output_dir = build_output_dir(pdf_path, args.output_dir, len(args.inputs))
        created_files = split_pdf(pdf_path, output_dir)
        print(f"{pdf_path}: wrote {len(created_files)} files to {output_dir}")
        if args.validate:
            problems = validate_outputs(created_files)
            for problem in problems:
                print(f"WARNING: {problem}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main(parser.parse_args()))
