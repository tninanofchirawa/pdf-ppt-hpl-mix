"""
Page Selector Orchestrator
==========================
Coordinates resolving input files, parsing page specifications, and saving extracted PDF.
"""

import os
from typing import Union, List
import pymupdf

from src.common.utils.file_utils import resolve_file_path, format_bytes
from src.pdf_tools.page_selector.range_parser import parse_page_spec
from src.pdf_tools.page_selector.extractor import extract_pages_to_file


def run_page_selection(input_pdf: str, output_pdf: str, page_spec: Union[str, List]) -> str:
    """
    Parses and extracts selected pages from source PDF to destination PDF.
    """
    resolved_pdf = resolve_file_path(input_pdf)
    doc = pymupdf.open(resolved_pdf)
    total_pages = len(doc)
    print(f"[*] Opened '{resolved_pdf}' ({total_pages} pages).")

    indices = parse_page_spec(page_spec, total_pages)
    if not indices:
        doc.close()
        raise ValueError(f"No valid pages found in specification: '{page_spec}'")

    natural_pages = [idx + 1 for idx in indices]
    print(f"[*] Extracting {len(indices)} pages: {natural_pages}")

    os.makedirs(os.path.dirname(os.path.abspath(output_pdf)), exist_ok=True)
    extract_pages_to_file(doc, indices, output_pdf)
    doc.close()

    new_size = os.path.getsize(output_pdf)
    print(f"[+] Saved extracted pages to '{output_pdf}' ({format_bytes(new_size)}).")
    return output_pdf
