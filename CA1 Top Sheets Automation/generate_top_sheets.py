"""Create one completed CA1 top sheet (DOCX and PDF) for every spreadsheet row.

Pass any XLSX workbook and DOCX template with --workbook and --template.
Requires Microsoft Word on Windows for the PDF-export step.
"""

from __future__ import annotations

import argparse
from concurrent.futures import ProcessPoolExecutor
import hashlib
import os
import random
import re
import sys
from pathlib import Path
from typing import Any

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.shared import Pt
from openpyxl import load_workbook


REQUIRED_COLUMNS = {
    "Name": "name",
    "UNI Roll number": "roll_number",
    "Mobile Number": "mobile_number",
    "Marks Obtained": "marks_obtained",
}

# Each label occurs once in the supplied template.  Keeping these in one place
# makes future template wording changes simple to handle.
TEMPLATE_FIELDS = {
    "name": "Name of the Student:",
    "roll_number": "Roll Number:",
    "mobile_number": "Mobile Number:",
    "marks_obtained": "Marks obtained:",
}

QUESTION_ONE_PARTS = ("1.a)", "1.b)", "1.c)", "1.d)", "1.e)", "1.f)", "1.g)")
LONG_QUESTIONS = ("2", "3", "4", "5", "6", "7")

# These are bundled with the standard Windows font collection on the computer
# used for this project.  Word falls back gracefully if a replacement machine
# does not have one, but installing the same font retains the intended look.
HANDWRITING_FONTS = (
    "Brush Script MT",
    "Freestyle Script",
    "Lucida Handwriting",
    "Mistral",
    "Viner Hand ITC",
)
EXAMINATION_DATE = "02-09-2026"


def as_text(value: Any) -> str:
    """Convert worksheet values without turning IDs into values such as 1.0."""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def clean_filename(value: str) -> str:
    """Make a Windows-safe, readable filename."""
    value = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", value).strip(". ")
    return value or "unnamed_student"


def parse_total_mark(value: str) -> int | None:
    """Return a numeric total, or None when the student was absent."""
    if value.casefold() == "ab":
        return None
    try:
        mark = int(value)
    except ValueError as error:
        raise ValueError(f"Marks Obtained must be a whole number from 0 to 25 or 'Ab'; got {value!r}.") from error
    if not 0 <= mark <= 25:
        raise ValueError(f"Marks Obtained must be from 0 to 25; got {mark}.")
    return mark


def distribute_marks(total: int, student_key: str) -> dict[str, int]:
    """Allocate a total within the CA1 attempt and per-question constraints.

    Question 1 contributes up to five 1-mark parts.  The remaining marks are
    spread across at most four of Questions 2--7, each capped at five marks.
    A roll-number seed gives each student a reproducible, varied distribution.
    """
    question_one_total = min(total, 5)
    remaining = total - question_one_total
    rng = random.Random(hashlib.sha256(student_key.encode("utf-8")).digest())
    allocation: dict[str, int] = {}

    question_one_choices = list(QUESTION_ONE_PARTS)
    rng.shuffle(question_one_choices)
    for question in question_one_choices[:question_one_total]:
        allocation[question] = 1

    # A positive score is assigned to up to four attempted long questions.
    long_question_choices = list(LONG_QUESTIONS)
    rng.shuffle(long_question_choices)
    attempted_count = min(4, remaining)
    attempted = long_question_choices[:attempted_count]
    for question in attempted:
        allocation[question] = 1
    remaining -= attempted_count

    # Add the rest one mark at a time, never exceeding the five-mark cap.
    while remaining:
        available = [question for question in attempted if allocation[question] < 5]
        if not available:  # Defensive guard: should be impossible for a /25 total.
            raise ValueError("Cannot allocate the requested mark within the CA1 question limits.")
        allocation[rng.choice(available)] += 1
        remaining -= 1

    if sum(allocation.values()) != total:
        raise AssertionError("Internal error: question-wise marks do not equal the total.")
    return allocation


def read_students(workbook_path: Path, sheet_name: str | None) -> list[dict[str, str]]:
    workbook = load_workbook(workbook_path, data_only=True, read_only=True)
    try:
        worksheet = workbook[sheet_name] if sheet_name else workbook.active
        # In read-only mode openpyxl represents blank cells as EmptyCell
        # instances, which do not have a ``column`` attribute. Enumerating the
        # row also works for normal cells and makes blank trailing columns safe.
        headers = {
            as_text(cell.value): column_number
            for column_number, cell in enumerate(
                next(worksheet.iter_rows(min_row=1, max_row=1)), start=1
            )
            if as_text(cell.value)
        }
        missing = [header for header in REQUIRED_COLUMNS if header not in headers]
        if missing:
            raise ValueError(f"Missing required column(s): {', '.join(missing)}")

        students: list[dict[str, str]] = []
        for row_number, row in enumerate(worksheet.iter_rows(min_row=2, values_only=True), start=2):
            student = {
                field: as_text(row[headers[header] - 1])
                for header, field in REQUIRED_COLUMNS.items()
            }
            if not any(student.values()):
                continue
            if not student["name"] or not student["roll_number"]:
                print(f"Skipping spreadsheet row {row_number}: name or roll number is blank.")
                continue
            students.append(student)
        return students
    finally:
        workbook.close()


def iter_template_paragraphs(document: Document):
    """Yield body and table paragraphs, including nested tables, exactly once."""
    # Keep references (rather than ``id(...)`` values): python-docx creates
    # short-lived wrapper objects while walking cells, and a recycled object ID
    # could otherwise make a later, real cell look like a duplicate.
    seen_cells: list[object] = []

    def yield_table_paragraphs(table: Table):
        for row in table.rows:
            for cell in row.cells:
                # Merged Word cells can be exposed more than once by python-docx.
                if any(cell._tc is seen_cell for seen_cell in seen_cells):
                    continue
                seen_cells.append(cell._tc)
                yield from cell.paragraphs
                for nested_table in cell.tables:
                    yield from yield_table_paragraphs(nested_table)

    yield from document.paragraphs
    for table in document.tables:
        yield from yield_table_paragraphs(table)


def replace_template_field(document: Document, label: str, value: str) -> None:
    """Replace a labelled field while retaining the label's existing formatting."""
    matches = [p for p in iter_template_paragraphs(document) if label in p.text]
    if len(matches) != 1:
        raise ValueError(
            f"Expected exactly one paragraph containing {label!r}; found {len(matches)}. "
            "The Word template may have changed."
        )

    paragraph = matches[0]
    # The supplied template keeps every label in one paragraph. Clearing then
    # recreating the paragraph avoids unreliable replacements across Word runs.
    for run in paragraph.runs:
        run.text = ""
    label_run = paragraph.add_run(label + " ")
    value_run = paragraph.add_run(value)
    # Use the label's formatting for the student value when Word has a direct
    # character style applied. (New runs otherwise use the paragraph style.)
    if len(paragraph.runs) >= 2:
        value_run.style = label_run.style


def set_cell_text(cell, value: str) -> None:
    """Set a marks-awarded cell without changing its paragraph alignment/style."""
    paragraph = cell.paragraphs[0]
    for run in paragraph.runs:
        run.text = ""
    paragraph.add_run(value)


def populate_marks_awarded(document: Document, total_mark: int | None, student_key: str) -> None:
    """Write the allocation into the template's Marks Awarded column."""
    if total_mark is None:
        # Absent students retain blank question-wise cells and show "Ab" in
        # the total field, as supplied in the workbook.
        return

    allocation = distribute_marks(total_mark, student_key)
    marks_table = next(
        (
            table
            for table in document.tables
            if table.rows
            and len(table.rows[0].cells) >= 3
            and table.rows[0].cells[0].text.strip() == "Q. No."
            and table.rows[0].cells[2].text.strip() == "Marks Awarded"
        ),
        None,
    )
    if marks_table is None:
        raise ValueError("Could not find the Marks Tabulation table in the Word template.")

    for row in marks_table.rows[1:]:
        question = row.cells[0].text.strip()
        if question in allocation:
            set_cell_text(row.cells[2], str(allocation[question]))


def insert_paragraph_before(paragraph) -> Paragraph:
    """Create a normal Word paragraph immediately before an existing one."""
    new_paragraph_xml = OxmlElement("w:p")
    paragraph._p.addprevious(new_paragraph_xml)
    return Paragraph(new_paragraph_xml, paragraph._parent)


def add_student_signature(document: Document, student_name: str, student_key: str) -> str:
    """Add a typed, handwriting-font signature just above the student line."""
    label = "Signature of the student with date"
    matches = [paragraph for paragraph in document.paragraphs if label in paragraph.text]
    if len(matches) != 1:
        raise ValueError(
            f"Expected exactly one paragraph containing {label!r}; found {len(matches)}. "
            "The Word template may have changed."
        )

    label_paragraph = matches[0]
    signature_paragraph = insert_paragraph_before(label_paragraph)
    signature_paragraph.alignment = label_paragraph.alignment or WD_ALIGN_PARAGRAPH.LEFT
    signature_paragraph.paragraph_format.left_indent = label_paragraph.paragraph_format.left_indent
    signature_paragraph.paragraph_format.space_after = Pt(0)

    font_rng = random.Random(hashlib.sha256((student_key + "-signature").encode("utf-8")).digest())
    font_name = font_rng.choice(HANDWRITING_FONTS)
    name_run = signature_paragraph.add_run(student_name)
    name_run.font.name = font_name
    name_run.font.size = Pt(16)
    # Explicitly set the East Asian font too, so Word does not replace the
    # selected face when opening the DOCX under another locale.
    name_run._element.rPr.rFonts.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia", font_name)

    date_run = signature_paragraph.add_run("\n" + EXAMINATION_DATE)
    date_run.font.size = Pt(9)
    return font_name


def create_docx(template_path: Path, student: dict[str, str], output_path: Path) -> None:
    document = Document(template_path)
    for field, label in TEMPLATE_FIELDS.items():
        replace_template_field(document, label, student[field])
    populate_marks_awarded(
        document,
        parse_total_mark(student["marks_obtained"]),
        student["roll_number"],
    )
    add_student_signature(document, student["name"], student["roll_number"])
    document.save(output_path)


def create_docx_job(job: tuple[str, dict[str, str], str]) -> str:
    """Process-pool worker: create one DOCX and return its filename stem."""
    template_path, student, output_path = job
    create_docx(Path(template_path), student, Path(output_path))
    return Path(output_path).stem


def export_pdf_with_word(docx_path: Path, pdf_path: Path) -> None:
    """Export through Word's native PDF engine, preserving the template layout."""
    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as error:
        raise RuntimeError("pywin32 is not installed. Run: py -m pip install -r requirements.txt") from error

    word = None
    document = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        document = word.Documents.Open(str(docx_path.resolve()))
        # 17 is Word's wdExportFormatPDF constant.
        document.ExportAsFixedFormat(str(pdf_path.resolve()), 17)
    except Exception as error:
        raise RuntimeError(
            "PDF export failed. Microsoft Word must be installed and able to open the template."
        ) from error
    finally:
        if document is not None:
            document.Close(False)
        if word is not None:
            word.Quit()


def main() -> int:
    here = Path(__file__).resolve().parent
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--workbook",
        type=Path,
        required=True,
        metavar="PATH",
        help="Path to the source XLSX workbook.",
    )
    parser.add_argument(
        "--template",
        type=Path,
        required=True,
        metavar="PATH",
        help="Path to the DOCX top-sheet template.",
    )
    parser.add_argument("--sheet", help="Worksheet name; defaults to the active worksheet.")
    parser.add_argument("--output", type=Path, default=here / "output")
    parser.add_argument(
        "--workers",
        type=int,
        default=min(4, os.cpu_count() or 1),
        help="Number of parallel DOCX-generation processes (default: up to 4).",
    )
    parser.add_argument("--keep-docx-only", action="store_true", help="Create DOCX files but skip PDF export.")
    args = parser.parse_args()

    for path, description in ((args.workbook, "workbook"), (args.template, "template")):
        if not path.is_file():
            parser.error(f"The {description} was not found: {path}")

    students = read_students(args.workbook, args.sheet)
    if not students:
        print("No student rows were found.")
        return 1
    if args.workers < 1:
        parser.error("--workers must be at least 1")

    docx_dir = args.output / "docx"
    pdf_dir = args.output / "pdf"
    docx_dir.mkdir(parents=True, exist_ok=True)
    if not args.keep_docx_only:
        pdf_dir.mkdir(parents=True, exist_ok=True)

    jobs = []
    for student in students:
        stem = clean_filename(f"{student['roll_number']}_{student['name']}")
        jobs.append((str(args.template), student, str(docx_dir / f"{stem}.docx")))

    print(f"Creating {len(students)} DOCX file(s) with {args.workers} worker(s)...")
    with ProcessPoolExecutor(max_workers=args.workers) as executor:
        stems = list(executor.map(create_docx_job, jobs))
    print("DOCX generation complete.")

    if not args.keep_docx_only:
        # Word's COM automation is intentionally single-process. Parallel Word
        # instances can leave orphaned WINWORD.EXE processes or lock files,
        # while parallel DOCX creation above provides the safe speed-up.
        print(f"Exporting {len(stems)} PDF file(s) through Microsoft Word...")
        for number, stem in enumerate(stems, start=1):
            export_pdf_with_word(docx_dir / f"{stem}.docx", pdf_dir / f"{stem}.pdf")
            print(f"[{number}/{len(stems)}] {stem}")

    print(f"Done. Files are in: {args.output.resolve()}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as error:
        print(f"ERROR: {error}", file=sys.stderr)
        raise SystemExit(1)
