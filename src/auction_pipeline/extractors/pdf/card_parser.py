"""
PDF Card Text Parser
====================
Extracts player name, rating points, matches played, batting and bowling statistics.
"""

import re
from typing import Dict, Any


KNOWN_LABELS = {
    'RATING', 'MATCHES PLAYED', 'BATTING STATS', 'BATTING STATS:',
    'BATTING AVG', 'RUNS SCORED', 'STRIKE RATE', 'BOWLING STATS',
    'BOWLING STATS:', 'BOWLING ECONOMY', 'WICKETS TAKEN', 'MATCHES PLAYED -'
}

IGNORE_CANDIDATES = {'PHOTO GOES HERE', 'PHOTO'}


def parse_player_card_text(text: str, page_idx: int) -> Dict[str, Any]:
    """
    Parses textual statistics and extracts the player's name and performance metrics.
    """
    lines = [l.strip().replace('\ufffd', '').strip() for l in text.split('\n') if l.strip()]

    # 1. Player Name
    name_candidates = []
    for line in lines:
        clean_l = re.sub(r'[\s\-\:\ufffd]+', ' ', line).strip().upper()
        if any(clean_l.startswith(k) or clean_l == k for k in KNOWN_LABELS):
            continue
        if re.match(r'^[\d\.\-\s]+$', line):
            continue
        if line.strip().upper() in IGNORE_CANDIDATES:
            continue
        name_candidates.append(line.strip())

    player_name = name_candidates[0] if name_candidates else f"Player_Page_{page_idx + 1}"

    # 2. Rating (5 - 10)
    rating_m = re.search(r'RATING\s*(\d+)', text)
    if not rating_m:
        for i, l in enumerate(lines):
            if l == 'RATING' and i + 1 < len(lines):
                rating_m = re.search(r'(\d+)', lines[i + 1])
                if rating_m:
                    break
    rating = rating_m.group(1) if rating_m else "-"

    # 3. Matches Played
    matches_m = re.search(r'MATCHES PLAYED[^\d]*(\d+)', text, re.I)
    if matches_m:
        matches = matches_m.group(1)
    elif 'Matches played -' in text:
        # Handles card variant (e.g., MS Dhoni)
        dm = re.search(r'RATING\s+\d+\s+(\d+)\s+BATTING', text)
        matches = dm.group(1) if dm else "-"
    else:
        matches = "-"

    # 4. Batting Stats
    bat_avg_m = re.search(r'BATTING AVG[^\d\-]*([\d\.\-]+)', text)
    runs_m = re.search(r'RUNS SCORED[^\d\-]*([\d\.\-]+)', text)
    sr_m = re.search(r'STRIKE RATE[^\d\-]*([\d\.\-]+)', text)

    # 5. Bowling Stats
    bowl_econ_m = re.search(r'BOWLING ECONOMY[^\d\-]*([\d\.\-]+)', text)
    wickets_m = re.search(r'WICKETS TAKEN[^\d\-]*([\d\.\-]+)', text)

    return {
        "Player_Name": player_name,
        "Rating_Points": rating,
        "Matches": matches,
        "Bat_Avg": bat_avg_m.group(1) if bat_avg_m else "-",
        "Runs": runs_m.group(1) if runs_m else "-",
        "Strike_Rate": sr_m.group(1) if sr_m else "-",
        "Bowl_Econ": bowl_econ_m.group(1) if bowl_econ_m else "-",
        "Wickets": wickets_m.group(1) if wickets_m else "-"
    }
