"""
Page Selector Extractor
=======================
Copies selected pages from a source PDF directly to a new output PDF.
"""

import pymupdf
from typing import List


def extract_pages_to_file(source_doc: pymupdf.Document, indices: List[int], output_pdf: str) -> None:
    """
    Directly inserts designated pages into a new PDF document.
    """
    out_doc = pymupdf.open()
    for idx in indices:
        out_doc.insert_pdf(source_doc, from_page=idx, to_page=idx)

    out_doc.save(output_pdf, garbage=4, deflate=True)
    out_doc.close()
