# This script merges multiple PDF files into a single PDF file.
# It uses the PyPDF2 library's PdfMerger to append PDFs in order and write the combined output.

from PyPDF2 import PdfMerger

merger = PdfMerger()
merger.append("C:/Personal/Directory/With/PDF1.pdf")
merger.append("C:/Personal/Directory/With/PDF2.pdf")
merger.append("C:/Personal/Directory/With/PDF3.pdf")


merger.write("C:/Personal/Directory/With/Merged.pdf")
merger.close()
