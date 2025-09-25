# This script unlocks a password-protected PDF, extracts specific data from it, and prints the results in a tabular format.
# It uses the pikepdf library to unlock the PDF and pdfplumber to extract tables

# importing necessary libraries
import pikepdf
import pdfplumber
import re
import os
import csv
import openpyxl

# Importing tabulate for better table formatting in the console
from tabulate import tabulate

# Step 1: Unlock the PDF
def unlock_pdf(input_path, output_path, password):
    with pikepdf.open(input_path, password=password) as pdf:
        pdf.save(output_path)

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
                    Ref_data = 'Ref. '
                else:
                    text_data = ''
                    value_date = ''
                    Ref_data = ''
                    
                    Row_value = result.split('\n')
                    for item in Row_value:                        
                        text_data = text_data + item.strip() + '\n' if item.strip() else text_data
                        if 'Value Dt' in item:
                            value_date_tmp = re.search(r"Value Dt (\d{2}/\d{2}/\d{4})", item)
                            value_date = value_date + value_date_tmp.group(1) + '\n' if value_date_tmp  else ''
                            text_data = text_data.replace('Value Dt '+ value_date_tmp.group(1), '')
                        if 'Ref ' in item:
                            ref_tmp = re.search(r"Ref\s+([^\n]+)", item)
                            Ref_data = Ref_data + ref_tmp.group(1) + '\n' if ref_tmp  else ''
                            text_data = text_data.replace('Ref '+ ref_tmp.group(1), '')                        
                        text_data = text_data.strip() + '\n'             

                new_row.extend([text_data, value_date, Ref_data])
                row_number += 1
            # Keep remaining columns
            new_row.extend(row[column_index + 1:])  
            new_table.append(new_row)
    return new_table

# Step 3: Append the cleaned data to a CSV file
def append_to_csv(file_path, data):
    file_exists = os.path.isfile(file_path)
    with open(file_path, mode='a', newline='\n') as file:
        writer = csv.writer(file)
        # if not file_exists:
        #     writer.writerow(headers)
        writer.writerows(data)


def append_to_excel(file_path, data):
    file_exists = os.path.isfile(file_path)

    if file_exists:
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook.active
    else:
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        # sheet.append(headers)


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

    workbook.save(file_path)

# Step 4: Process all PDFs in a directory
def process_pdfs_in_directory(directory):
    for root, _, files in os.walk(directory):
        for file in files:
            if file.lower().endswith('.pdf'):
                pdf_path = os.path.join(root, file)
                unlocked_pdf_path = os.path.join(root, "t." + file)
                print(f"Processing: {pdf_path}")
                unlock_pdf(pdf_path, unlocked_pdf_path, password)
                extract_and_print_tables(unlocked_pdf_path)
                # Clean up the temporary unlocked file
                os.remove(unlocked_pdf_path)


# Step 5: Extract and print tables from page 2 onward
def extract_and_print_tables(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages[1:], start=2):  # Skip page 1
            print(f"\n--- Page {i} ---")
            tables = page.extract_tables()
            for table_index, table in enumerate(tables):
                print(f"\n #### Table {table_index + 1}: (Before cleaning) ####")
                print(tabulate(table, headers="firstrow", tablefmt="grid"))                
                cleaned_table = extract_parts_from_column(table, column_index=1)                
                print(f"\n #### Table {table_index + 1}: (After cleaning) ####")
                print(tabulate(cleaned_table, headers="firstrow", tablefmt="grid"))                
                append_to_excel(op_excel_path, cleaned_table)
                
# Replace with your actual file path
pdf_directory = "C:/Personal/Directory/With/PDFs"       
# Temporary unlocked file
# unlocked_pdf = "C:/Personal/Directory/With/PDFs/unlocked.pdf"
# outut file path
op_excel_path = "C:/Personal/Directory/With/PDFs/output.xlsx"
# Replace with your actual password
password = "pdf-file-password"

process_pdfs_in_directory(pdf_directory)

# unlock_pdf(locked_pdf, unlocked_pdf, password)
# extract_and_print_tables(unlocked_pdf)
# os.remove(unlocked_pdf)  # Clean up the temporary unlocked file
