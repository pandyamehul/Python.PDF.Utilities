# compress_pdf.py
from pypdf import PdfReader, PdfWriter

def compress_pdf(input_path, output_path):
    """
    Compress a PDF file by removing duplication and compressing content streams.
    
    Args:
        input_path: Path to input PDF file
        output_path: Path to save compressed PDF file
    """
    reader = PdfReader(input_path)
    writer = PdfWriter()
    
    # Copy all pages to writer
    for page in reader.pages:
        writer.add_page(page)
    
    # Compress content streams after pages are added
    for page in writer.pages:
        page.compress_content_streams()
    
    # Write compressed PDF with additional compression
    with open(output_path, "wb") as output_file:
        writer.write(output_file)
    
    print(f"Compressed PDF saved to: {output_path}")

# Usage example
if __name__ == "__main__":
    input_pdf = "scan-Merged.pdf"
    output_pdf = "compressed_scan-Merged.pdf"
    compress_pdf(input_pdf, output_pdf)