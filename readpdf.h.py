"""This script unlocks password-protected PDFs, extracts transaction-style
tables from them, and prints the results in a tabular format.

It uses the pikepdf library to unlock the PDF and pdfplumber to extract
tables. Cleaned tables are appended to an Excel workbook. All paths and the
password are configurable via command-line arguments instead of being
hardcoded, so the script works with any directory/file location.

Usage:
    python readpdf.h.py -d /path/to/pdf/dir -o output.xlsx -p <password>

If -p/--password is omitted, the script prompts for it interactively without
echoing it to the console.
"""

import argparse
import getpass
import io
import os
import re
import sys

import openpyxl
import pdfplumber
import pikepdf
from tabulate import tabulate


def parse_args(argv=None):
    """Parse command-line arguments for the PDF extraction script."""
    parser = argparse.ArgumentParser(
        description="Unlock PDFs in a directory, extract tables, and export to Excel."
    )
    parser.add_argument(
        "-d",
        "--directory",
        required=True,
        help="Directory containing the password-protected PDF files to process.",
    )
    parser.add_argument(
        "-o",
        "--output",
        required=True,
        help="Path of the Excel file where extracted tables will be appended.",
    )
    parser.add_argument(
        "-p",
        "--password",
        default=None,
        help="Password used to unlock the PDF files. Prompted for if omitted.",
    )
    return parser.parse_args(argv)


# Step 1: Unlock the PDF
def unlock_pdf(input_path, password):
    """Remove password protection from a PDF and return the unlocked bytes in memory."""
    buffer = io.BytesIO()
    with pikepdf.open(input_path, password=password) as pdf:
        pdf.save(buffer)
    buffer.seek(0)
    return buffer


# Step 2: Extract 'Text data', 'Ref. Date', and 'Trn value' from column 2
def extract_parts_from_column(table, column_index=1):
    new_table = []
    row_number = 1
    for row in table:
        if row is not None:
            new_row = row[:column_index]  # Keep columns before the target
            if len(row) > column_index and row[column_index]:
                cell_value = row[column_index]

                pattern = r"(Ref\s+\w+)\n"
                replacement = r"\1|"
                result = re.sub(pattern, replacement, cell_value)

                # Replace \n only when it follows 'Date <date> Ref'
                pattern = r"(Value Dt\s+\d{2}/\d{2}/\d{4})\n"
                replacement = r"\1 Ref #NA|"
                result = re.sub(pattern, replacement, result)

                result = result.replace('\n', ' ')
                # Replace '|' with newline for better readability
                result = result.replace('|', '\n')

                if row_number == 1:
                    text_data = 'Narration Text'
                    value_date = 'Value Dt'
                    ref_data = 'Ref. '
                else:
                    text_data = ''
                    value_date = ''
                    ref_data = ''

                    row_values = result.split('\n')
                    for item in row_values:
                        text_data = text_data + item.strip() + '\n' if item.strip() else text_data
                        if 'Value Dt' in item:
                            value_date_tmp = re.search(r"Value Dt (\d{2}/\d{2}/\d{4})", item)
                            value_date = value_date + value_date_tmp.group(1) + '\n' if value_date_tmp else ''
                            text_data = text_data.replace('Value Dt ' + value_date_tmp.group(1), '')
                        if 'Ref ' in item:
                            ref_tmp = re.search(r"Ref\s+([^\n]+)", item)
                            ref_data = ref_data + ref_tmp.group(1) + '\n' if ref_tmp else ''
                            text_data = text_data.replace('Ref ' + ref_tmp.group(1), '')
                        text_data = text_data.strip() + '\n'

                new_row.extend([text_data, value_date, ref_data])
                row_number += 1
            # Keep remaining columns
            new_row.extend(row[column_index + 1:])
            new_table.append(new_row)
    return new_table


# Step 3: Append the cleaned data to an Excel workbook, splitting multi-line cells into rows
def append_to_excel(file_path, data):
    if os.path.isfile(file_path):
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook.active
    else:
        workbook = openpyxl.Workbook()
        sheet = workbook.active

    for row in data:
        # Determine the maximum number of lines in any cell of the row
        max_lines = max(len(str(cell).split('\n')) for cell in row)
        # Create new rows for each line
        for i in range(max_lines):
            new_row = []
            for cell in row:
                cell_lines = str(cell).split('\n')
                new_row.append(cell_lines[i] if i < len(cell_lines) else '')
            sheet.append(new_row)
            # openpyxl would store a leading '=' as a formula; keep PDF text as plain text
            for cell in sheet[sheet.max_row]:
                if isinstance(cell.value, str) and cell.value.startswith('='):
                    cell.data_type = 's'

    workbook.save(file_path)


# Step 4: Extract and print tables from page 2 onward
def extract_and_print_tables(pdf_source, output_excel_path):
    with pdfplumber.open(pdf_source) as pdf:
        for i, page in enumerate(pdf.pages[1:], start=2):  # Skip page 1
            print(f"\n--- Page {i} ---")
            tables = page.extract_tables()
            for table_index, table in enumerate(tables):
                print(f"\n #### Table {table_index + 1}: (Before cleaning) ####")
                print(tabulate(table, headers="firstrow", tablefmt="grid"))
                cleaned_table = extract_parts_from_column(table, column_index=1)
                print(f"\n #### Table {table_index + 1}: (After cleaning) ####")
                print(tabulate(cleaned_table, headers="firstrow", tablefmt="grid"))
                append_to_excel(output_excel_path, cleaned_table)


# Step 5: Process all PDFs in a directory
def process_pdfs_in_directory(directory, password, output_excel_path):
    for root, _, files in os.walk(directory):
        for file in files:
            if file.lower().endswith('.pdf'):
                pdf_path = os.path.join(root, file)
                print(f"Processing: {pdf_path}")
                try:
                    unlocked_pdf = unlock_pdf(pdf_path, password)
                    extract_and_print_tables(unlocked_pdf, output_excel_path)
                except pikepdf.PasswordError:
                    print(f"Error: incorrect password for {pdf_path}, skipping.", file=sys.stderr)


def main(argv=None):
    args = parse_args(argv)

    if not os.path.isdir(args.directory):
        print(f"Error: directory not found: {args.directory}", file=sys.stderr)
        return 1

    password = args.password or getpass.getpass("Enter PDF password: ")

    process_pdfs_in_directory(args.directory, password, args.output)
    return 0


if __name__ == "__main__":
    sys.exit(main())
