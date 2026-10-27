"""
Phase 4: Presentation Generator Orchestrator
============================================
Coordinates generating the polished HPL 2.0 Mega Player Auction presentation deck.
"""

import os
from typing import Optional
from generate_hpl2_deck import generate_hpl2_presentation
from src.common.utils.file_utils import resolve_file_path


def run_presentation_generation(excel_path: Optional[str] = None,
                                output_pptx: str = "HPL_2_Auction_Deck.pptx") -> str:
    """
    Builds the HPL 2.0 auction presentation deck from the consolidated dataset.
    """
    if excel_path is None:
        excel_path = "outputs/Master_Cricket_Auction.xlsx"

    resolved_excel = resolve_file_path(excel_path)
    return generate_hpl2_presentation(players_data=resolved_excel, output_path=output_pptx)
