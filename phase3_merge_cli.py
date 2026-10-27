"""
Phase 3: Dataset Consolidation Entrypoint
========================================
Merges current and historical auction records into Master_Cricket_Auction.xlsx.
"""

import sys
import os
import argparse

sys.path.insert(0, os.path.abspath(os.path.dirname(__file__)))

from src.common.config.settings import DEFAULT_CURRENT_EXCEL, DEFAULT_PREVIOUS_EXCEL, DEFAULT_MASTER_EXCEL
from src.auction_pipeline.merger import run_merger


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Phase 3: Merge and deduplicate player datasets")
    parser.add_argument("--current", default=DEFAULT_CURRENT_EXCEL, help="Current extracted Excel file")
    parser.add_argument("--previous", default=DEFAULT_PREVIOUS_EXCEL, help="Previous extracted Excel file")
    parser.add_argument("--output", default=DEFAULT_MASTER_EXCEL, help="Output master Excel file")
    parser.add_argument("--auto", action="store_true", help="Run automatically without interactive prompts")
    args = parser.parse_args()

    run_merger(current_excel=args.current,
               previous_excel=args.previous,
               output_excel=args.output,
               auto_mode=args.auto)
