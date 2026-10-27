"""
Page Range Parser
=================
Parses natural 1-based page specs into zero-based document indices with bounds checking.
"""

from typing import List, Union


def parse_page_spec(spec: Union[str, List, tuple], total_pages: int) -> List[int]:
    """
    Parses natural 1-based page specifications into a list of valid 0-based indices.
    Examples:
      '1-4, 85-87, 182-185' -> [0, 1, 2, 3, 84, 85, 86, 181, 182, 183, 184]
    """
    selected_indices = []

    if isinstance(spec, (list, tuple)):
        raw_tokens = [str(item).strip() for item in spec]
    elif isinstance(spec, str):
        raw_tokens = [t.strip() for t in spec.split(",") if t.strip()]
    else:
        raw_tokens = [str(spec).strip()]

    for token in raw_tokens:
        if "-" in token:
            parts = token.split("-", 1)
            try:
                start = int(parts[0].strip())
                end = int(parts[1].strip())
            except ValueError:
                print(f"[!] Warning: Skipping malformed range '{token}'")
                continue

            if start > end:
                print(f"[!] Swapping inverted range '{token}' (start {start} > end {end})")
                start, end = end, start

            for p in range(start, end + 1):
                idx = p - 1
                if 0 <= idx < total_pages:
                    if idx not in selected_indices:
                        selected_indices.append(idx)
                else:
                    print(f"[!] Warning: Page {p} out of range (Document has {total_pages} pages)")
        else:
            try:
                page_num = int(token)
                idx = page_num - 1
                if 0 <= idx < total_pages:
                    if idx not in selected_indices:
                        selected_indices.append(idx)
                else:
                    print(f"[!] Warning: Page {page_num} out of range (Document has {total_pages} pages)")
            except ValueError:
                print(f"[!] Warning: Skipping invalid token '{token}'")

    return selected_indices
