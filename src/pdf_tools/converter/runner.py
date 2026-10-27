"""
PDF Converter Orchestrator
==========================
Coordinates rasterization and slide insertion to convert a PDF into a PowerPoint deck.
"""

import os
import gc
from typing import Optional, List
import pymupdf

from src.common.utils.file_utils import resolve_file_path, format_bytes
from src.pdf_tools.converter.rasterizer import rasterize_page
from src.pdf_tools.converter.pptx_builder import initialize_presentation, append_page_slide


def run_conversion(pdf_path: str,
                   output_pptx: str,
                   page_indices: Optional[List[int]] = None,
                   dpi: int = 120,
                   img_format: str = "jpeg",
                   jpeg_quality: int = 85,
                   aspect_ratio: str = "match") -> str:
    """
    Executes the full PDF-to-PPTX conversion pipeline.
    """
    resolved_pdf = resolve_file_path(pdf_path)
    doc = pymupdf.open(resolved_pdf)
    total_pages = len(doc)

    if page_indices is None:
        target_indices = list(range(total_pages))
    else:
        target_indices = [idx for idx in page_indices if 0 <= idx < total_pages]

    if not target_indices:
        doc.close()
        raise ValueError("No valid pages selected for conversion.")

    prs = initialize_presentation(doc[target_indices[0]].rect, aspect_ratio=aspect_ratio)
    print(f"[*] Slide size: {prs.slide_width.inches:.2f}\" x {prs.slide_height.inches:.2f}\" (Mode: {aspect_ratio})")
    print(f"[*] Converting {len(target_indices)} pages at {dpi} DPI ({img_format.upper()})...")

    for step_idx, page_idx in enumerate(target_indices):
        page = doc[page_idx]
        stream = rasterize_page(page, dpi=dpi, img_format=img_format, jpeg_quality=jpeg_quality)
        append_page_slide(prs, stream)
        stream.close()

        if (step_idx + 1) % 25 == 0 or (step_idx + 1) == len(target_indices):
            print(f"    Progress: [{step_idx + 1}/{len(target_indices)}] slides converted...")
            gc.collect()

    doc.close()
    os.makedirs(os.path.dirname(os.path.abspath(output_pptx)), exist_ok=True)
    prs.save(output_pptx)
    file_size = os.path.getsize(output_pptx)
    print(f"[+] Presentation saved successfully: '{output_pptx}' ({format_bytes(file_size)}).")
    return output_pptx
