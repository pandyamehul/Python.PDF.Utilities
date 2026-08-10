# Python PDF Utility Project to automate smaller task with PDF Files

## Merge PDF Files

- Simple python script that merges multiple PDF files into a single PDF file using the PyPDF2 library.
- You can specify the paths of the PDF files you want to merge and the output file path where the merged PDF will be saved.
- Make sure to install the PyPDF2 library using pip if you haven't already:
  
  ```pwsh
  pip install PyPDF2
  ```

- Run the script from the command line, passing the input PDF paths and the desired output path. It will create a merged PDF file containing all the specified PDFs in the order they were added:

  ```pwsh
  python mergepdf.py -o Merged.pdf PDF1.pdf PDF2.pdf PDF3.pdf
  ```

Ref. file - mergepdf.py

## Read/Extract Data from PDF Files

- Python script that unlocks password-protected PDFs in a directory, extracts transaction-style tables from them, and appends the cleaned data to an Excel workbook.
- Make sure to install the required libraries using pip if you haven't already:

  ```pwsh
  pip install pikepdf pdfplumber tabulate openpyxl
  ```

- Run the script from the command line, passing the PDF directory, output Excel file path, and password. If the password is omitted, you will be prompted for it:

  ```pwsh
  python readpdf.h.py -d /path/to/pdf/dir -o output.xlsx -p <password>
  ```

Ref. file - readpdf.h.py

## Compare PDF Files

- This section is under development. The goal is to create a script that compares two PDF files and highlights the differences between them.

Ref. files - compare.pdf.1.py, compare.pdf.2.py
