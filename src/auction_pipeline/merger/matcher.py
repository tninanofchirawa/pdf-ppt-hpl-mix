"""
Player Matcher & Enricher
=========================
Identifies candidate matches between current and historical auction records using fuzzy string comparison.
"""

from typing import Tuple, Optional
import pandas as pd
from src.common.utils.text_utils import name_similarity


def find_best_candidate(curr_name: str, unmatched_df: pd.DataFrame, threshold: float = 0.80) -> Tuple[Optional[int], float]:
    """
    Finds the highest-scoring candidate row in the previous dataset.
    Returns: (best_index, best_score)
    """
    if "Player_Name" not in unmatched_df.columns or unmatched_df.empty:
        return None, 0.0

    best_idx = None
    best_score = 0.0

    for idx, row in unmatched_df.iterrows():
        prev_name = str(row.get("Player_Name", ""))
        score = name_similarity(curr_name, prev_name)
        if score > best_score:
            best_score = score
            best_idx = idx

    if best_score >= threshold:
        return best_idx, best_score
    return None, best_score


def enrich_player_record(curr_rec: dict, prev_row: pd.Series) -> dict:
    """
    Merges historical attributes (Base Price, Age, Country) into current player profile.
    """
    enriched = dict(curr_rec)
    enriched["Base_Price"] = prev_row.get("Base_Price", "-")
    enriched["Age"] = prev_row.get("Age", "-")
    if pd.isna(enriched.get("Country")) or enriched.get("Country") == "-":
        enriched["Country"] = prev_row.get("Country", "-")
    enriched["Merged_Status"] = "Merged_Enriched"
    return enriched
