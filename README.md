# XLSX to Templated Docs

Generate individual Word documents (and optionally PDFs) by merging rows from an Excel spreadsheet into a Word template with mail-merge-style placeholders.

## How It Works

1. **Read** an Excel file (`.xlsx`) and iterate over each row in a specified sheet.
2. **Replace** `«Placeholder»` tokens in a Word template (`.docx`) with the corresponding cell values for that row.
3. **Save** one populated document per row, named using key fields from the data (e.g. PI name, nominee, nomination type).
4. **Convert** (optional) the generated `.docx` files to PDF via LibreOffice.

## Prerequisites

- Python 3.10+
- [LibreOffice](https://www.libreoffice.org/) (only required for PDF conversion)

## Setup

### Local

```bash
pip install -r requirements.txt
```

### Dev Container

A dev container configuration is included. Open the project in VS Code / Cursor with the Dev Containers extension and it will automatically install Python dependencies and LibreOffice.

## Input Files

Place the following files in the project root (they are git-ignored):

| File | Description |
|---|---|
| `Cleaned EOIs.xlsx` | Excel workbook with source data (reads from the **"2. Filtered-Complete EOIs"** sheet) |
| `Round 1 and Data.docx` | Word template containing `«Placeholder»` tokens |

## Placeholder Mapping

The template uses mail-merge-style `«Placeholder»` tokens that map to Excel columns:

| Template Placeholder | Excel Column |
|---|---|
| `«PI_Name_Ucalgary_System»` | PI Name (Ucalgary System) |
| `«PI_Faculty_Ucalgary_system»` | PI Faculty (Ucalgary system) |
| `«PI_Department»` | PI Department |
| `«PI_Department_Ucalgary_system»` | PI Department (Ucalgary system) |
| `«Training_Environment»` | Training Environment |
| `«Project_Title»` | Project Title |
| `«Nomination_Type»` | Nomination Type |
| `«Applicable_Impact_Priority_Areas»` | Applicable Impact+ Priority Areas |
| `«Applicable_UCalgary_Transdisciplinary_Ar»` | Applicable UCalgary Transdisciplinary Areas of Focus |
| `«Most_Relevant_TriAgency»` | Most Relevant Tri-Agency |
| `«Abstract»` | Abstract |
| `«International_PostDoc_Count_2024»` | International Post-Doc Count (2024) |
| `«National_PostDoc_Count_2024»` | National Post-Doc Count (2024) |
| `«Grad_Student_Count_2024»` | Grad Student Count (2024) |

## Usage

### 1. Upload input files

The Excel data file and Word template are **not** included in the repository. You need to upload them into the project root before running the script:

| File to upload | Description |
|---|---|
| `Cleaned EOIs.xlsx` | Excel workbook with source data |
| `Round 1 and Data.docx` | Word template containing `«Placeholder»` tokens |

**In GitHub Codespaces**, you can upload files by dragging them into the Explorer sidebar, or by using the terminal:

```bash
# From the terminal inside Codespaces, you can also use the GUI:
# Right-click the file explorer → Upload...
```

### 2. Run the script

```bash
# Inspect Excel columns and template placeholders (no files generated)
python generate_documents.py --inspect

# Preview what would be generated without creating files
python generate_documents.py --dry-run

# Generate Word documents
python generate_documents.py

# Generate Word documents and convert to PDF
python generate_documents.py --pdf

# Convert existing .docx files in output/ to PDF (no generation)
python generate_documents.py --pdf-only
```

### 3. Download the output

Generated files are saved to `output/` (`.docx`) and `output_pdf/` (`.pdf`).

**In GitHub Codespaces**, download the results by:

- Right-clicking the `output/` or `output_pdf/` folder in the Explorer sidebar and selecting **Download...**.
- Or downloading individual files by right-clicking them and selecting **Download...**.

Each output file is named as `[PI Name] [Nominee Name] [Nomination Type].docx`. Accented characters are transliterated to ASCII for filename safety.

## Project Structure

```
.
├── .devcontainer/
│   └── devcontainer.json   # Dev container config (Python 3.12 + LibreOffice)
├── .gitignore
├── generate_documents.py   # Main script
├── requirements.txt        # Python dependencies
└── README.md
```
