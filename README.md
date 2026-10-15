# Python PDF Utility Project

A small collection of Python scripts that automate common PDF tasks: merging, extracting table data, compressing, and comparing PDFs.

## Scripts at a glance

| Script | What it does | Input | Output | Configured by |
| --- | --- | --- | --- | --- |
| `mergepdf.py` | Merges PDFs in the order given | PDF paths on the command line | One merged PDF | Command-line arguments |
| `readpdf.h.py` | Unlocks PDFs and extracts transaction tables | A folder of password-protected PDFs | Rows appended to an `.xlsx` file | Command-line arguments |
| `compress_pdf.py` | Recompresses content streams | One PDF | One smaller PDF | Edit the `__main__` block |
| `compress_pdf_advanced.py` | Re-encodes large images as JPEG | One PDF | One smaller PDF | Edit the `__main__` block |
| `compare_pdf.txt-summary.py` | Lists the changed text lines between two PDFs | Two PDFs | Text file of changes | Edit the bottom of the script |
| `compare_pdf_html-summary.py` | Builds a side-by-side HTML diff of two PDFs | `compare.pdf.file.1.pdf` and `compare.pdf.file.2.pdf` | `pdf_diff_report.html` | Edit the bottom of the script |

## Setup (virtual environment)

Run these commands from the project folder (PowerShell on Windows):

```pwsh
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

On macOS/Linux, use `source .venv/bin/activate` instead of the activate script.

`requirements.txt` installs every library used by the scripts. Each section below lists the ones it needs.

## Merge PDF Files (`mergepdf.py`)

Merges PDFs into one file, in the order you list them. Uses **pypdf**.

```pwsh
python mergepdf.py -o Merged.pdf PDF1.pdf PDF2.pdf PDF3.pdf
```

- `-o/--output` (required): path of the merged PDF. Missing folders are created.
- Positional arguments: input PDFs, in the order they should appear.
- If any input file is missing, the script lists the missing files and exits without writing output.

## Read and Extract Table Data (`readpdf.h.py`)

Unlocks each password-protected PDF in a folder (including subfolders), extracts transaction-style tables from page 2 onward, prints them, and appends the cleaned rows to an Excel workbook. Uses **pikepdf**, **pdfplumber**, **tabulate**, and **openpyxl**.

```pwsh
python readpdf.h.py -d "C:\path\to\pdf\folder" -o output.xlsx
```

- `-d/--directory` (required): folder containing the PDFs.
- `-o/--output` (required): Excel file to create or append to.
- `-p/--password` (optional): PDF password. If omitted, the script prompts for it without echoing. Prefer the prompt, because a password passed on the command line is visible in shell history and process listings.

Notes:

- Unlocked data is processed in memory and is never written to disk.
- PDFs with a wrong password are reported and skipped.
- Multi-line cells are split into separate Excel rows.
- Cell text that starts with `=` is stored as plain text, so it cannot run as an Excel formula.

## Compress PDF Files

Both scripts take their input and output paths from the `__main__` block at the bottom of the file. Edit those lines, then run:

```pwsh
python compress_pdf.py
python compress_pdf_advanced.py
```

### `compress_pdf.py`

Recompresses page content streams and writes a new PDF. Uses **pypdf**. Quick and lossless, but usually gives a modest size reduction, mainly for PDFs with uncompressed content.

### `compress_pdf_advanced.py`

Re-encodes images as JPEG for a larger size reduction. Uses **pikepdf** and **Pillow**.

- `image_quality` (default 85, range 0-100): lower values give smaller files and more visible artifacts. The `__main__` block uses 25.
- Images wider or taller than 1920 pixels are scaled down to fit.
- Images with transparency masks (`/SMask`, `/Mask`, `/ImageMask`) or decode arrays are left unchanged.
- The summary at the end reports how many images were compressed and how many were skipped.

## Compare PDF Files

Both scripts extract the text of two PDFs and compare it. They compare **text only**, so scanned PDFs without a text layer will show no differences. Both use **PyMuPDF** (`fitz`).

Both scripts have example paths at the bottom of the file. Edit them before running.

### `compare_pdf.txt-summary.py`: text summary

Writes only the lines that changed between the two PDFs. Each line is prefixed with `+` (added in the second PDF) or `-` (removed from the first).

```pwsh
python compare_pdf.txt-summary.py
```

The default output file is `pdf_diff_summary.txt`. The script currently calls `compare_pdfs()` with the two example PDF paths at the bottom of the file.

### `compare_pdf_html-summary.py`: HTML report

Builds a side-by-side HTML diff, with changed lines highlighted, and saves it as `pdf_diff_report.html` in the current folder.

```pwsh
python compare_pdf_html-summary.py
```

By default, the script reads `compare.pdf.file.1.pdf` and `compare.pdf.file.2.pdf` from the current folder. The first PDF appears on the left side of the report.

## Security notes

- Do not commit PDFs, passwords, or generated outputs. `.gitignore` excludes `*.pdf*` and `*.html*`, and `.venv/` is ignored.
- See [SECURITY.md](SECURITY.md) for how to report a vulnerability.
