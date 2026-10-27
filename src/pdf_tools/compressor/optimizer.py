"""
PDF Stream Optimizer
====================
Optimizes PDF files through garbage collection, unreferenced object removal, and stream deflation.
"""

import pymupdf


def optimize_pdf_file(input_path: str, output_path: str) -> None:
    """
    Saves an optimized copy of the PDF, stripping unused objects and compressing streams.
    """
    doc = pymupdf.open(input_path)
    doc.save(
        output_path,
        garbage=4,           # Remove unreferenced & dead objects
        deflate=True,        # Deflate content streams
        deflate_images=True, # Recompress embedded image streams with ZIP
        deflate_fonts=True,  # Deflate font streams
        clean=True           # Sanitize stream syntax
    )
    doc.close()
