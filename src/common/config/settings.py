"""
Application Settings & Constants
================================
Central configuration for directories, inputs, outputs, and theme styling.
"""

import os
from pptx.dml.color import RGBColor

# Base Directories
WORKSPACE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", ".."))
SRC_DIR = os.path.join(WORKSPACE_DIR, "src")
INPUTS_DIR = os.path.join(WORKSPACE_DIR, "inputs")
OUTPUTS_DIR = os.path.join(WORKSPACE_DIR, "outputs")
DOCS_DIR = os.path.join(WORKSPACE_DIR, "docs")

# Ensure required directory structure exists
os.makedirs(INPUTS_DIR, exist_ok=True)
os.makedirs(OUTPUTS_DIR, exist_ok=True)
os.makedirs(DOCS_DIR, exist_ok=True)

# Default Input Paths
DEFAULT_PDF_INPUT = os.path.join(INPUTS_DIR, "Cricket_Auction.pdf")
DEFAULT_PPTX_INPUT = os.path.join(INPUTS_DIR, "IPL AUCTION.pptx")

# Default Output Paths
DEFAULT_PPTX_DIRECT_OUTPUT = os.path.join(OUTPUTS_DIR, "Cricket Auction.pptx")
DEFAULT_SELECTED_PDF_OUTPUT = os.path.join(OUTPUTS_DIR, "Cricket_Auction_Selected.pdf")
DEFAULT_COMPRESSED_PDF_OUTPUT = os.path.join(OUTPUTS_DIR, "compressed.pdf")

DEFAULT_CURRENT_EXCEL = os.path.join(OUTPUTS_DIR, "Extracted_Current.xlsx")
DEFAULT_PREVIOUS_EXCEL = os.path.join(OUTPUTS_DIR, "Extracted_Previous.xlsx")
DEFAULT_MASTER_EXCEL = os.path.join(OUTPUTS_DIR, "Master_Cricket_Auction.xlsx")
DEFAULT_FINAL_PPTX = os.path.join(WORKSPACE_DIR, "HPL_2_Auction_Deck.pptx")

DEFAULT_CURRENT_IMG_DIR = os.path.join(OUTPUTS_DIR, "images_current")
DEFAULT_PREVIOUS_IMG_DIR = os.path.join(OUTPUTS_DIR, "images_previous")

# Slide Dimensions (Inches)
SLIDE_WIDTH_16_9 = 13.333
SLIDE_HEIGHT_16_9 = 7.5
SLIDE_WIDTH_4_3 = 10.0
SLIDE_HEIGHT_4_3 = 7.5

# HPL 2.0 Color Palette
THEME_COLOR_BG = RGBColor(11, 15, 25)           # Deep Navy Dark (#0B0F19)
THEME_COLOR_CARD_BG = RGBColor(22, 31, 48)      # Dark Slate Blue (#161F30)
THEME_COLOR_CARD_BORDER = RGBColor(39, 53, 79)  # Subtle Border (#27354F)
THEME_COLOR_GOLD = RGBColor(250, 204, 21)       # Gold Accent (#FACC15)
THEME_COLOR_PRICE_GREEN = RGBColor(34, 197, 94) # Electric Green (#22C55E)
THEME_COLOR_SECONDARY = RGBColor(156, 163, 175) # Light Gray (#9CA3AF)
THEME_COLOR_WHITE = RGBColor(248, 249, 250)     # White (#F8F9FA)
