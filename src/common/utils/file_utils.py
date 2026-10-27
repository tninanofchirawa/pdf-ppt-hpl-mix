"""
File & Filesystem Utilities
===========================
Helper functions for resolving file paths, formatting sizes, and sanitizing names.
"""

import os
import re


def resolve_file_path(path: str, fallback_candidates: list = None) -> str:
    """
    Resolves a file path by checking:
      1. Exact path
      2. Alternative with spaces / underscores
      3. Inside 'inputs/' directory
      4. Inside 'outputs/' directory
      5. Inside workspace root
    Raises FileNotFoundError if none exist.
    """
    if os.path.exists(path):
        return os.path.abspath(path)

    # Calculate search directories
    workspace_root = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", ".."))
    inputs_dir = os.path.join(workspace_root, "inputs")
    outputs_dir = os.path.join(workspace_root, "outputs")

    base_name = os.path.basename(path)
    base_name_alt = base_name.replace("_", " ") if "_" in base_name else base_name.replace(" ", "_")

    search_candidates = [
        # Check space/underscore in original directory
        path.replace("_", " ") if "_" in path else path.replace(" ", "_"),
        # Check inside inputs/
        os.path.join(inputs_dir, base_name),
        os.path.join(inputs_dir, base_name_alt),
        # Check inside outputs/
        os.path.join(outputs_dir, base_name),
        os.path.join(outputs_dir, base_name_alt),
        # Check inside workspace root
        os.path.join(workspace_root, base_name),
        os.path.join(workspace_root, base_name_alt),
    ]

    if fallback_candidates:
        search_candidates.extend(fallback_candidates)

    for cand in search_candidates:
        if os.path.exists(cand):
            return os.path.abspath(cand)

    raise FileNotFoundError(
        f"Cannot locate file '{path}'. Checked candidates in inputs/, outputs/, and root: {search_candidates}"
    )


def sanitize_filename(name: str) -> str:
    """
    Converts a name or title into a filesystem-safe filename string.
    """
    clean = re.sub(r'[^a-zA-Z0-9_\-\s]', '', str(name)).strip()
    return re.sub(r'\s+', '_', clean).lower()


def format_bytes(size_bytes: int) -> str:
    """
    Formats byte size to human-readable string (B, KB, MB, GB).
    """
    for unit in ['B', 'KB', 'MB', 'GB']:
        if size_bytes < 1024.0:
            return f"{size_bytes:3.2f} {unit}"
        size_bytes /= 1024.0
    return f"{size_bytes:.2f} TB"
