"""This script merges multiple PDF files into a single PDF file.

It uses the PyPDF2 library's PdfMerger to append PDFs in order and write the
combined output. Input and output paths are provided via command-line
arguments so the script works with any file location, instead of hardcoded
paths.

Usage:
    python mergepdf.py -o Merged.pdf PDF1.pdf PDF2.pdf PDF3.pdf
"""

import argparse
import sys
from pathlib import Path

from PyPDF2 import PdfMerger


def parse_args(argv=None):
    """Parse command-line arguments for the PDF merge script."""
    parser = argparse.ArgumentParser(
        description="Merge multiple PDF files into a single PDF file."
    )
    parser.add_argument(
        "input_files",
        nargs="+",
        help="Paths of the PDF files to merge, in the order they should appear.",
    )
    parser.add_argument(
        "-o",
        "--output",
        required=True,
        help="Path of the merged output PDF file.",
    )
    return parser.parse_args(argv)


def merge_pdfs(input_paths, output_path):
    """Merge the given PDF files, in order, into a single output PDF file.

    Raises:
        FileNotFoundError: If any of the input files does not exist.
    """
    missing_files = [str(path) for path in input_paths if not Path(path).is_file()]
    if missing_files:
        raise FileNotFoundError(f"Input file(s) not found: {', '.join(missing_files)}")

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with PdfMerger() as merger:
        for input_path in input_paths:
            merger.append(str(input_path))
        merger.write(str(output_path))


def main(argv=None):
    args = parse_args(argv)
    try:
        merge_pdfs(args.input_files, args.output)
    except (FileNotFoundError, OSError) as error:
        print(f"Error: {error}", file=sys.stderr)
        return 1

    print(f"Merged {len(args.input_files)} file(s) into: {args.output}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
