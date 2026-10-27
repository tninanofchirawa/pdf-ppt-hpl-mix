"""
PDF Portrait & Nationality Extractor
====================================
Extracts high-resolution player portraits and detects overseas flight icons.
"""

import os
from typing import Tuple
import pymupdf

from src.common.utils.file_utils import sanitize_filename


def extract_player_portrait(doc: pymupdf.Document,
                            page: pymupdf.Page,
                            player_name: str,
                            img_dir: str) -> Tuple[str, bool]:
    """
    Extracts the player's portrait from the left card area.
    Detects if the player is Overseas based on the flight icon in the top right.
    Returns: (saved_image_path, is_overseas)
    """
    saved_img_path = ""
    is_overseas = False

    for img_info in page.get_images():
        xref = img_info[0]
        base_img = doc.extract_image(xref)
        w, h = base_img["width"], base_img["height"]
        rects = page.get_image_rects(xref)
        r = rects[0] if rects else None

        # Overseas indicator: small airplane icon in the upper-right corner
        if w < 200 and h < 200 and r and r.x0 > 500:
            is_overseas = True
        # Player portrait: large photo in the upper-left card section
        elif w > 200 and r and r.x0 < 250 and r.y0 < 300:
            safe_name = sanitize_filename(player_name)
            saved_img_path = os.path.join(img_dir, f"{safe_name}.png")
            try:
                with open(saved_img_path, "wb") as f_img:
                    f_img.write(base_img["image"])
            except Exception as e:
                print(f"[!] Warning: Could not write image for '{player_name}': {e}")

    return saved_img_path, is_overseas
