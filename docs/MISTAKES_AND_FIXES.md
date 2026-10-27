# Comprehensive Autopsy: Mistakes & Solutions

This document details every bug, design limitation, and failure mode found in the original codebase, along with the implemented fixes.

---

## 1. `run.py` (PDF to PPT Converter)

* **Bug 1 (Crash on Launch):** Hardcoded `pdf_path = "Cricket Auction.pdf"`. The actual file was `Cricket_Auction.pdf` (with an underscore). Threw `FileNotFoundError`.
  * *Fix:* Implemented `resolve_file_path` which automatically checks `inputs/`, `outputs/`, and root directories for filename variants (spaces vs underscores).
* **Bug 2 (Aspect Ratio Distortion):** Fixed dimensions `13.333" x 7.5"` (16:9 widescreen) distorted the 4:3 PDF (`768 x 576` pts).
  * *Fix:* Added `--aspect-ratio match` (default) which reads the source PDF's physical dimensions and sets matching PowerPoint slide dimensions.
* **Bug 3 (High Memory & File Size):** Uncompressed PNG bytes caused high memory usage and a 135 MB file.
  * *Fix:* Added JPEG compression option (default 85% quality, reducing file size by ~70%) with explicit garbage collection.

---

## 2. `phase1_extract_pdf.py` (Current PDF Extraction)

* **Bug 1 (Failure to Extract Player Names):** Regex assumed the player's name preceded `"RATING"`. In the PDF, `"RATING"` appears first, followed by the rating integer, with the name appearing below or at the bottom of the card.
  * *Fix:* Built a candidate line filter that discards known stat labels and numeric lines to isolate the exact player name line.
* **Bug 2 (Broken Fallback):** `if l == "RATING" and i > 0:` was always False because `"RATING"` was at line index 0. All 249 players defaulted to `Player_Page_X`.
  * *Fix:* Completely overhauled parsing to accurately extract all 249 player names with 0 missing records.
* **Bug 3 (Broken Category Detection):** Tested for text `"BATTERS"` or `"BOWLERS"`, but divider slides (pages 1, 84, 181) were raster images with zero text (`""`), assigning `"General"` to every player.
  * *Fix:* Identified document section boundaries: Pages 2–83 (Batters & WKs), Pages 85–180 (Bowlers), Pages 182–252 (All Rounders), with WKs distinguished by the `(WK)` tag.
* **Bug 4 (Placeholder Text):** On Page 161, `'PHOTO GOES HERE'` was extracted as a player name.
  * *Fix:* Added placeholder exclusion filter, accurately identifying `NUWAN THUSHARA`.
* **Bug 5 (Overseas vs Domestic):** No mechanism to distinguish overseas players.
  * *Fix:* Added coordinate-based detection for the top-right airplane flight icon to label players as **Overseas** or **India**.

---

## 3. `phase2_extract_pptx.py` (Previous PPTX Extraction)

* **Bug 1 (Trapezoid Badges Misidentified):** In `IPL AUCTION.pptx`, player names, base prices, and ratings are raster graphic image cutouts inside GroupShapes, while `BATTER`, `BOWLER`, `INDIA`, and `AGE` are plain text boxes. The script assigned `"BATTER"` and `"BOWLER"` as player names for 100+ slides.
  * *Fix:* Added recursive shape walker to un-nest GroupShapes, extract candidate portrait photos, and prevent category labels from polluting player name fields.

---

## 4. `phase3_merge_cli.py` (Data Merging)

* **Bug 1 (0% Match Score):** Merging `Player_Page_X` against `BATTER` resulted in 0% similarity.
  * *Fix:* With Phase 1 and 2 clean data, fuzzy matching scores accurately. Added an `--auto` flag for non-interactive automated pipeline runs.

---

## 5. `phase4_generate_pptx.py` (Presentation Generator)

* **Bug 1 (Failure & Missing Assets):** Script had `# This was an utter failure` at line 1 and depended on non-existent files `player_card_bg.png` and `divider_bg.png`.
  * *Fix:* Built a standalone **Dark Stadium Auction Theme** using native PowerPoint vector shapes (navy `#0F172A`, slate containers, gold accents, stats tables, and scaled portrait frames). No external background dependencies required.
