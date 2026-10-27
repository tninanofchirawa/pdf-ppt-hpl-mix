"""
Phase 2: PPTX Extractor Orchestrator
====================================
Coordinates PPTX slide traversal, picture extraction, and Excel output.
"""

import os
import pandas as pd
from pptx import Presentation

from src.common.utils.file_utils import resolve_file_path, sanitize_filename
from src.auction_pipeline.extractors.pptx.shape_walker import walk_shapes_recursive, extract_candidate_portraits
from src.auction_pipeline.extractors.pptx.card_parser import parse_pptx_slide_text


def run_pptx_extraction(pptx_path: str, output_excel: str, img_dir: str) -> pd.DataFrame:
    """
    Extracts player slides and attributes from a PowerPoint deck.
    """
    resolved_pptx = resolve_file_path(pptx_path)
    os.makedirs(img_dir, exist_ok=True)
    os.makedirs(os.path.dirname(os.path.abspath(output_excel)), exist_ok=True)

    prs = Presentation(resolved_pptx)
    records = []
    current_category = "General"
    print(f"[*] Extracting slides from '{resolved_pptx}' ({len(prs.slides)} slides)...")

    for slide_idx, slide in enumerate(prs.slides):
        all_shapes = list(walk_shapes_recursive(slide.shapes))
        text_chunks = []

        for shape in all_shapes:
            if shape.has_text_frame:
                for p in shape.text_frame.paragraphs:
                    txt = p.text.strip()
                    if txt:
                        text_chunks.append(txt)

        combined = " ".join(text_chunks).upper()

        # Update category section
        if "THE BATTERS" in combined:
            current_category = "Batter"
            continue
        elif "THE BOWLERS" in combined:
            current_category = "Bowler"
            continue
        elif "THE ALL ROUNDERS" in combined or "ALL-ROUNDERS" in combined:
            current_category = "All Rounder"
            continue
        elif "WICKET KEEPERS" in combined or "WICKET" in combined:
            current_category = "Wicket Keeper"
            continue

        portraits = extract_candidate_portraits(all_shapes)
        if "AGE" not in combined and not portraits:
            continue

        card_info = parse_pptx_slide_text(text_chunks, current_category)
        player_id = f"Previous_Player_Slide_{slide_idx + 1}"

        img_path = ""
        if portraits:
            best_img = portraits[0]
            safe_id = sanitize_filename(player_id)
            img_path = os.path.join(img_dir, f"{safe_id}.{best_img[2]}")
            with open(img_path, "wb") as f_img:
                f_img.write(best_img[1])

        records.append({
            "Slide_Num": slide_idx + 1,
            "Player_ID": player_id,
            "Category": card_info["Category"],
            "Age": card_info["Age"],
            "Country": card_info["Country"],
            "Image_Path": img_path,
            "Source": "Previous_PPTX"
        })

    df = pd.DataFrame(records)
    df.to_excel(output_excel, index=False)
    print(f"[+] Phase 2 Complete: {len(df)} slides extracted into '{output_excel}'.")
    return df
