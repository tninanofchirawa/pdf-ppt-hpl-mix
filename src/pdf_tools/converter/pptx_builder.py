"""
PowerPoint Deck Builder
=======================
Creates presentation slides and places rasterized pages with accurate aspect ratios.
"""

from pptx import Presentation
from pptx.util import Inches


def initialize_presentation(sample_page_rect, aspect_ratio: str = "match") -> Presentation:
    """
    Initializes a PowerPoint presentation with slide dimensions matching the requested aspect ratio.
    """
    prs = Presentation()

    pdf_w_pts = sample_page_rect.width
    pdf_h_pts = sample_page_rect.height

    if aspect_ratio == "16:9":
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
    elif aspect_ratio == "4:3":
        prs.slide_width = Inches(10.0)
        prs.slide_height = Inches(7.5)
    else:  # 'match' exact PDF physical dimensions
        prs.slide_width = Inches(pdf_w_pts / 72.0)
        prs.slide_height = Inches(pdf_h_pts / 72.0)

    return prs


def append_page_slide(prs: Presentation, img_stream) -> None:
    """
    Creates a blank slide and attaches the raster image full bleed.
    """
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)
    slide.shapes.add_picture(
        img_stream,
        0,
        0,
        width=prs.slide_width,
        height=prs.slide_height
    )
