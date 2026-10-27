"""
PDF Category Detector
=====================
Determines player auction category based on page section ranges and role suffixes.
"""

def detect_category(page_idx: int, player_name: str) -> str:
    """
    Identifies whether the player belongs to Batter, Bowler, All Rounder, or Wicket Keeper.
    Section bounds in Cricket_Auction.pdf:
      - Pages 2-83  : Batters (and WKs if '(WK)' present)
      - Pages 85-180: Bowlers
      - Pages 182-252: All Rounders
    """
    if "(WK)" in player_name.upper():
        return "Wicket Keeper"

    if page_idx <= 83:
        return "Batter"
    elif page_idx <= 180:
        return "Bowler"
    else:
        return "All Rounder"
