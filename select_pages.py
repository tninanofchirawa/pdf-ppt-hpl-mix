"""
Select Pages Entrypoint
=======================
Extracts specific pages or page ranges from a source PDF.
"""

import sys
import os
import argparse

sys.path.insert(0, os.path.abspath(os.path.dirname(__file__)))

from src.common.config.settings import DEFAULT_PDF_INPUT, DEFAULT_SELECTED_PDF_OUTPUT
from src.pdf_tools.page_selector import run_page_selection, parse_page_spec


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Extract selected pages from a PDF")
    parser.add_argument("-i", "--input", default=DEFAULT_PDF_INPUT, help="Input PDF path")
    parser.add_argument("-o", "--output", default=DEFAULT_SELECTED_PDF_OUTPUT, help="Output PDF path")
    parser.add_argument("-p", "--pages", default="1-4, 85-87, 182-185", help="Page range specification (e.g. '1-4, 85-87')")
    args = parser.parse_args()

    run_page_selection(input_pdf=args.input, output_pdf=args.output, page_spec=args.pages)
