
import fitz  # PyMuPDF
import difflib

def extract_text_from_pdf(file_path):
    """Extract text from all pages of a PDF file."""
    doc = fitz.open(file_path)
    text = ''
    for page in doc:
        text += page.get_text()
    return text

def get_diff_summary(text1, text2):
    """Return only the changed lines between two texts."""
    diff = difflib.ndiff(text1.splitlines(), text2.splitlines())
    changes = [line for line in diff if line.startswith('+ ') or line.startswith('- ')]
    return '\n'.join(changes)

def compare_pdfs(pdf1_path, pdf2_path, output_path='pdf_diff_summary.txt'):
    """Compare two PDFs and save the summary of changes to a file."""
    text1 = extract_text_from_pdf(pdf1_path)
    text2 = extract_text_from_pdf(pdf2_path)
    summary = get_diff_summary(text1, text2)
    
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(summary)
    
    print(f"Summary of changes saved to {output_path}")

# Example usage
compare_pdfs('C:/Personal/Directory/With/PDFs/pdf-file-1.pdf', 'C:/Personal/Directory/With/PDFs/pdf-file-2.pdf')