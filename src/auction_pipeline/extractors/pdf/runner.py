"""
Phase 1: PDF Extractor Orchestrator
===================================
Coordinates parsing, portrait extraction, categorization, and Excel export.
"""

import os
import pandas as pd
import pymupdf

from src.common.utils.file_utils import resolve_file_path
from src.auction_pipeline.extractors.pdf.card_parser import parse_player_card_text
from src.auction_pipeline.extractors.pdf.category_detector import detect_category
from src.auction_pipeline.extractors.pdf.portrait_extractor import extract_player_portrait


def run_pdf_extraction(pdf_path: str, output_excel: str, img_dir: str) -> pd.DataFrame:
    """
    Parses player cards from the PDF auction file and produces structured Excel output.
    """
    resolved_pdf = resolve_file_path(pdf_path)
    os.makedirs(img_dir, exist_ok=True)
    os.makedirs(os.path.dirname(os.path.abspath(output_excel)), exist_ok=True)

    doc = pymupdf.open(resolved_pdf)
    records = []
    print(f"[*] Extracting players from '{resolved_pdf}' ({len(doc)} pages)...")

    for page_idx in range(len(doc)):
        page = doc[page_idx]
        text = page.get_text()

        # Filter out cover and divider slides
        if 'BATTING STATS' not in text and 'BOWLING STATS' not in text:
            continue

        # 1. Parse text statistics
        card_data = parse_player_card_text(text, page_idx)
        player_name = card_data["Player_Name"]

        # 2. Determine category
        category = detect_category(page_idx, player_name)

        # 3. Extract portrait & detect Overseas / India
        img_path, is_overseas = extract_player_portrait(doc, page, player_name, img_dir)
        country = "Overseas" if is_overseas else "India"

        record = {
            "Page_Num": page_idx + 1,
            "Player_Name": player_name,
            "Category": category,
            "Country": country,
            "Rating_Points": card_data["Rating_Points"],
            "Matches": card_data["Matches"],
            "Bat_Avg": card_data["Bat_Avg"],
            "Runs": card_data["Runs"],
            "Strike_Rate": card_data["Strike_Rate"],
            "Bowl_Econ": card_data["Bowl_Econ"],
            "Wickets": card_data["Wickets"],
            "Image_Path": img_path,
            "Source": "Current_PDF"
        }
        records.append(record)

    doc.close()

    df = pd.DataFrame(records)
    df.to_excel(output_excel, index=False)
    print(f"[+] Phase 1 Complete: {len(df)} players extracted into '{output_excel}'.")
    return df
