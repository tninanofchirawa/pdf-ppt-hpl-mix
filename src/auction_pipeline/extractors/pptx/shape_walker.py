"""
PPTX Shape Walker
=================
Recursively traverses slide shape trees, un-nesting GroupShapes and extracting pictures.
"""

from typing import Generator, List, Tuple


def walk_shapes_recursive(shapes) -> Generator:
    """
    Recursively yields all shapes, penetrating any nested GroupShapes.
    """
    for shape in shapes:
        yield shape
        if hasattr(shape, "shapes"):
            yield from walk_shapes_recursive(shape.shapes)


def extract_candidate_portraits(all_shapes) -> List[Tuple[int, bytes, str]]:
    """
    Finds and ranks candidate player portraits from slide shapes.
    Returns list of tuples: (pixel_area, blob_bytes, extension)
    """
    candidates = []
    for shape in all_shapes:
        if shape.shape_type == 13 or hasattr(shape, "image"):
            try:
                img = shape.image
                w, h = img.size
                # Discard small icons, tiny banners, or large backgrounds
                if 200 < w < 2000 and 200 < h < 2000:
                    candidates.append((w * h, img.blob, img.ext))
            except Exception:
                pass

    candidates.sort(key=lambda x: x[0], reverse=True)
    return candidates
