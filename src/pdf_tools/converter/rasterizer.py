"""
PDF Page Rasterizer
===================
Converts PDF vector pages into high-performance compressed raster image bytes.
"""

import io
import pymupdf


def rasterize_page(page: pymupdf.Page, dpi: int = 120, img_format: str = "jpeg", jpeg_quality: int = 85) -> io.BytesIO:
    """
    Renders a PyMuPDF Page into an in-memory byte stream.
    Supports high-speed JPEG compression to reduce presentation file size.
    """
    pix = page.get_pixmap(dpi=dpi)

    if img_format.lower() in ["jpg", "jpeg"]:
        img_bytes = pix.tobytes("jpeg", jpg_quality=jpeg_quality)
    else:
        img_bytes = pix.tobytes("png")

    stream = io.BytesIO(img_bytes)
    pix = None  # Explicit memory release
    return stream
