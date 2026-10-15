import difflib
import fitz  # pip install PyMuPDF

# Function to extract text from a PDF file
def extract_text_from_pdf(file_path):
    doc = fitz.open(file_path)
    text = ''
    for page in doc:
        text += page.get_text()
    return text

# Generate HTML diff
def generate_html_diff(text1, text2, output_html='pdf_diff_report.html'):
    diff = difflib.HtmlDiff().make_file(
        text1.splitlines(),
        text2.splitlines(),
        fromdesc='PDF Document 1',
        todesc='PDF Document 2'
    )
    with open(output_html, 'w', encoding='utf-8') as f:
        f.write(diff)

# Usage
text1 = extract_text_from_pdf('compare.pdf.file.1.pdf')
text2 = extract_text_from_pdf('compare.pdf.file.2.pdf')
generate_html_diff(text1, text2)

