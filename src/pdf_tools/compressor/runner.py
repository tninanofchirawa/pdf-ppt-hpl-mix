"""
PDF Compressor Orchestrator
===========================
Coordinates PDF compression and produces comprehensive space-savings metrics.
"""

import os
from src.common.utils.file_utils import resolve_file_path, format_bytes
from src.pdf_tools.compressor.optimizer import optimize_pdf_file


def run_compression(input_pdf: str, output_pdf: str) -> str:
    """
    Compresses input PDF and prints detailed before/after compression metrics.
    """
    resolved_pdf = resolve_file_path(input_pdf)
    original_size = os.path.getsize(resolved_pdf)
    print(f"[*] Opening '{resolved_pdf}' ({format_bytes(original_size)})...")

    os.makedirs(os.path.dirname(os.path.abspath(output_pdf)), exist_ok=True)
    optimize_pdf_file(resolved_pdf, output_pdf)

    new_size = os.path.getsize(output_pdf)
    saved_bytes = original_size - new_size
    reduction_pct = (saved_bytes / original_size) * 100 if original_size > 0 else 0

    print(f"[+] Compressed file saved: '{output_pdf}'")
    print(f"    Original Size  : {format_bytes(original_size)}")
    print(f"    Compressed Size: {format_bytes(new_size)}")
    print(f"    Space Saved    : {format_bytes(saved_bytes)} ({reduction_pct:.1f}% reduction)")

    return output_pdf
