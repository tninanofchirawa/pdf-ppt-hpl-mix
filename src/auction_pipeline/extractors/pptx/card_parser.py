"""
PPTX Card Text Parser
=====================
Extracts age, country, and category attributes from slide text frames.
"""

import re
from typing import Dict, Any, List


def parse_pptx_slide_text(text_chunks: List[str], current_section_cat: str) -> Dict[str, Any]:
    """
    Parses structured text attributes from previous year presentation slide.
    """
    combined_text = "\n".join(text_chunks)

    age_m = re.search(r"AGE\s*:\s*(\d+)", combined_text, re.I)
    country_m = re.search(
        r"\b(INDIA|AUSTRALIA|ENGLAND|SOUTH AFRICA|NEW ZEALAND|AFGHANISTAN|WEST INDIES|WINDIES|SRI LANKA|SRILANKA|BANGLADESH)\b",
        combined_text, re.I
    )
    category_m = re.search(
        r"\b(BATTER|BOWLER|ALL ROUNDER|ALL-ROUNDER|ROUNDER|WK\+BATTER|WK\s*\+\s*BATTER|WICKET KEEPER)\b",
        combined_text, re.I
    )

    category = current_section_cat
    if category_m:
        c_str = category_m.group(1).upper()
        if "WK" in c_str or "WICKET" in c_str:
            category = "Wicket Keeper"
        elif "BOWL" in c_str:
            category = "Bowler"
        elif "ROUNDER" in c_str:
            category = "All Rounder"
        elif "BATTER" in c_str:
            category = "Batter"

    return {
        "Age": age_m.group(1) if age_m else "-",
        "Country": country_m.group(1).upper() if country_m else "-",
        "Category": category
    }
