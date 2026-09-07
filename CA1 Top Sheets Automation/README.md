# CA1 Top Sheets Automation

Generate one completed Word top sheet and PDF per student from any Excel workbook and compatible DOCX template.

The workbook and template are supplied at runtime, so their filenames and locations are entirely up to the user. Original input files are never changed.

## Features

- Reads student details and total marks from an XLSX workbook.
- Fills the matching fields in a DOCX top-sheet template.
- Allocates each numeric total across the question-wise `Marks Awarded` column without exceeding the mark for a question.
- Adds a typed student signature in one of five handwriting-style fonts, with the examination date beneath it.
- Creates files named `Rollnumber_Name.docx` and `Rollnumber_Name.pdf`.
- Generates DOCX files in parallel, then uses Microsoft Word for faithful PDF conversion.

## Requirements

- Windows
- Python 3.10 or later
- Microsoft Word (required only for PDF export)

Install the Python dependencies:

```powershell
py -m pip install -r requirements.txt
```

## Quick start

After placing or cloning this project on a Windows computer, open PowerShell in the project folder and run:

```powershell
py .\generate_top_sheets.py
```

Two standard Windows file pickers will open in sequence:

1. Choose the student Excel workbook (`.xlsx`).
2. Choose the Word top-sheet template (`.docx`).

The generated files are written to `output\docx` and `output\pdf`. Input filenames do not need to match any predefined names.

## Input requirements

### Excel workbook

The selected worksheet must have these exact headers in its first row:

| Header | Used for |
| --- | --- |
| `Name` | Student name and typed signature |
| `UNI Roll number` | Roll number and output filename |
| `Mobile Number` | Mobile Number field |
| `Marks Obtained` | Total score and question-wise allocation |

`Marks Obtained` must be a whole number from 0 to 25 or `Ab`.

### Word template

The template must contain these labels exactly once:

| Template label | Filled with |
| --- | --- |
| `Name of the Student:` | `Name` |
| `Roll Number:` | `UNI Roll number` |
| `Mobile Number:` | `Mobile Number` |
| `Marks obtained:` | `Marks Obtained` |
| `Signature of the student with date` | Typed signature and examination date are inserted above this label |

The question-allocation feature expects a table headed `Q. No.` and `Marks Awarded`, with Question 1 parts (`1.a)` through `1.g)`) and Questions 2 through 7.

## Command-line options (optional)

For repeatable or scripted runs, provide the paths directly instead of using the file pickers:

```powershell
py .\generate_top_sheets.py --workbook "C:\path\to\students.xlsx" --template "C:\path\to\top-sheet-template.docx"
```

By default, generated files are saved alongside the script:

```
output/
  docx/
    34900725001_Durjoy pandit.docx
  pdf/
    34900725001_Durjoy pandit.pdf
```

Useful options:

```powershell
# Use a named worksheet
py .\generate_top_sheets.py --workbook ".\students.xlsx" --template ".\template.docx" --sheet "Sheet1"

# Choose an output directory and six DOCX-generation workers
py .\generate_top_sheets.py --workbook ".\students.xlsx" --template ".\template.docx" --output ".\generated" --workers 6

# Create DOCX files only; do not export PDFs
py .\generate_top_sheets.py --workbook ".\students.xlsx" --template ".\template.docx" --keep-docx-only
```

## Mark allocation

For each numeric total, the generated `Marks Awarded` entries are verified to sum to `Marks Obtained`:

- Up to 5 marks are distributed as 1 mark each among the seven Question 1 parts.
- The balance is distributed across up to four of Questions 2--7, with no question receiving more than 5 marks.
- `Ab` leaves question-wise marks blank.

The allocation and signature font are varied but repeatable for a given roll number.

## Notes

- PDF export is intentionally sequential. Running multiple Microsoft Word automation instances in parallel can cause file locks and orphaned Word processes.
- The five signature fonts are common Windows fonts: Brush Script MT, Freestyle Script, Lucida Handwriting, Mistral, and Viner Hand ITC. Word will substitute a font if one is missing.
- Input workbooks, templates, and generated files are excluded by `.gitignore` so student data is not committed accidentally.
