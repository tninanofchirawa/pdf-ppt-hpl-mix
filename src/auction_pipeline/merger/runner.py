"""
Phase 3: Player Dataset Merger Orchestrator
===========================================
Consolidates extracted current and previous auction datasets into Master_Cricket_Auction.xlsx.
"""

import os
import pandas as pd

from src.common.utils.file_utils import resolve_file_path
from src.auction_pipeline.merger.matcher import find_best_candidate, enrich_player_record


def run_merger(current_excel: str,
               previous_excel: str,
               output_excel: str,
               auto_mode: bool = False) -> pd.DataFrame:
    """
    Consolidates player datasets into a master Excel file.
    """
    resolved_curr = resolve_file_path(current_excel)
    df_curr = pd.read_excel(resolved_curr)

    df_prev = pd.DataFrame()
    if os.path.exists(previous_excel):
        try:
            resolved_prev = resolve_file_path(previous_excel)
            df_prev = pd.read_excel(resolved_prev)
        except Exception:
            pass

    merged_records = []
    unmatched_prev = df_prev.copy() if not df_prev.empty else pd.DataFrame()

    print("=" * 60)
    print("CONSOLIDATING AUCTION DATASETS")
    print("=" * 60)
    print(f"[*] Current records : {len(df_curr)}")
    print(f"[*] Previous records: {len(df_prev)}")

    for _, curr_row in df_curr.iterrows():
        curr_rec = curr_row.to_dict()
        curr_name = str(curr_rec.get("Player_Name", ""))

        best_idx, best_score = find_best_candidate(curr_name, unmatched_prev)

        if best_idx is not None:
            p_row = unmatched_prev.loc[best_idx]
            if auto_mode:
                choice = "D"
            else:
                print("\n" + "-" * 50)
                print(f"[MATCH FOUND] {curr_name} <-> {p_row.get('Player_Name')} ({int(best_score*100)}%)")
                print("Options: [A] Keep Current Only  [B] Use Previous  [C] Keep Both  [D] Merge Enriched")
                choice = input("Enter choice (A/B/C/D) [Default D]: ").strip().upper() or "D"

            if choice == "A":
                curr_rec["Merged_Status"] = "Current_Only"
                merged_records.append(curr_rec)
                unmatched_prev = unmatched_prev.drop(best_idx)
            elif choice == "B":
                prev_rec = p_row.to_dict()
                prev_rec["Merged_Status"] = "Previous_Only"
                merged_records.append(prev_rec)
                unmatched_prev = unmatched_prev.drop(best_idx)
            elif choice == "C":
                curr_rec["Merged_Status"] = "Current_Unmerged"
                merged_records.append(curr_rec)
            else:  # Default "D"
                enriched = enrich_player_record(curr_rec, p_row)
                merged_records.append(enriched)
                unmatched_prev = unmatched_prev.drop(best_idx)
        else:
            curr_rec.setdefault("Base_Price", "-")
            curr_rec.setdefault("Age", "-")
            curr_rec["Merged_Status"] = "Current_Only"
            merged_records.append(curr_rec)

    # Append historical previous records if available and requested
    if not unmatched_prev.empty and "Player_Name" in unmatched_prev.columns:
        append_prev = True if auto_mode else (
            input(f"\nAppend {len(unmatched_prev)} unmatched previous players? (Y/N) [Default N]: ").strip().upper() == "Y"
        )
        if append_prev:
            for _, rem_row in unmatched_prev.iterrows():
                r = rem_row.to_dict()
                r["Merged_Status"] = "Previous_Historical"
                merged_records.append(r)

    master_df = pd.DataFrame(merged_records)
    os.makedirs(os.path.dirname(os.path.abspath(output_excel)), exist_ok=True)
    master_df.to_excel(output_excel, index=False)
    print(f"\n[+] Master dataset successfully saved ({len(master_df)} rows) to '{output_excel}'.")
    return master_df
