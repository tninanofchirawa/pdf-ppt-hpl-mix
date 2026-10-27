"""
Phase 2: PPTX Slide Extraction Entrypoint
=========================================
Extracts player slides, candidate photos, and attributes from previous PPTX.
"""

import sys
import os
import argparse

sys.path.insert(0, os.path.abspath(os.path.dirname(__file__)))

from src.common.config.settings import DEFAULT_PPTX_INPUT, DEFAULT_PREVIOUS_EXCEL, DEFAULT_PREVIOUS_IMG_DIR
from src.auction_pipeline.extractors.pptx import run_pptx_extraction


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Phase 2: Extract data from previous PPTX")
    parser.add_argument("--pptx", default=DEFAULT_PPTX_INPUT, help="Path to input PPTX file")
    parser.add_argument("--output", default=DEFAULT_PREVIOUS_EXCEL, help="Path to output Excel file")
    parser.add_argument("--img-dir", default=DEFAULT_PREVIOUS_IMG_DIR, help="Directory to save extracted images")
    args = parser.parse_args()

    run_pptx_extraction(pptx_path=args.pptx, output_excel=args.output, img_dir=args.img_dir)
