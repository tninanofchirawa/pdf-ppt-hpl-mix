"""
PDF Compressor Entrypoint
=========================
Compresses and optimizes PDF file size.
"""

import sys
import os
import argparse

sys.path.insert(0, os.path.abspath(os.path.dirname(__file__)))

from src.common.config.settings import DEFAULT_PDF_INPUT, DEFAULT_COMPRESSED_PDF_OUTPUT
from src.pdf_tools.compressor import run_compression


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Compress and optimize PDF files")
    parser.add_argument("-i", "--input", default=DEFAULT_PDF_INPUT, help="Input PDF path")
    parser.add_argument("-o", "--output", default=DEFAULT_COMPRESSED_PDF_OUTPUT, help="Output compressed PDF path")
    args = parser.parse_args()

    run_compression(input_pdf=args.input, output_pdf=args.output)
