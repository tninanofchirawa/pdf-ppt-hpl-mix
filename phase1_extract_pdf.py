"""
Phase 1: PDF Player Card Extraction Entrypoint
==============================================
Extracts player cards, stats, overseas classification, and portraits from PDF.
"""

import sys
import os
import argparse

sys.path.insert(0, os.path.abspath(os.path.dirname(__file__)))

from src.common.config.settings import DEFAULT_PDF_INPUT, DEFAULT_CURRENT_EXCEL, DEFAULT_CURRENT_IMG_DIR
from src.auction_pipeline.extractors.pdf import run_pdf_extraction


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Phase 1: Extract player stats & photos from PDF")
    parser.add_argument("--pdf", default=DEFAULT_PDF_INPUT, help="Path to input PDF file")
    parser.add_argument("--output", default=DEFAULT_CURRENT_EXCEL, help="Path to output Excel file")
    parser.add_argument("--img-dir", default=DEFAULT_CURRENT_IMG_DIR, help="Directory to save extracted player photos")
    args = parser.parse_args()

    run_pdf_extraction(pdf_path=args.pdf, output_excel=args.output, img_dir=args.img_dir)
