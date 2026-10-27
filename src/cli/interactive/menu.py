"""
Interactive Terminal Menu
=========================
Renders structured terminal menus with clear user options and guided workflows.
"""

import sys
import os

from src.common.config.settings import (
    DEFAULT_PDF_INPUT, DEFAULT_PPTX_INPUT, DEFAULT_PPTX_DIRECT_OUTPUT,
    DEFAULT_SELECTED_PDF_OUTPUT, DEFAULT_COMPRESSED_PDF_OUTPUT,
    DEFAULT_CURRENT_EXCEL, DEFAULT_PREVIOUS_EXCEL, DEFAULT_MASTER_EXCEL,
    DEFAULT_FINAL_PPTX, DEFAULT_CURRENT_IMG_DIR, DEFAULT_PREVIOUS_IMG_DIR
)
from src.common.utils.file_utils import format_bytes
from src.cli.interactive.prompts import ask_string, ask_choice, ask_int

from src.pdf_tools.converter import run_conversion
from src.pdf_tools.page_selector import run_page_selection, parse_page_spec
from src.pdf_tools.compressor import run_compression
from src.auction_pipeline.extractors.pdf import run_pdf_extraction
from src.auction_pipeline.extractors.pptx import run_pptx_extraction
from src.auction_pipeline.merger import run_merger
from src.auction_pipeline.generator import run_presentation_generation


def print_banner():
    print("\n" + "=" * 64)
    print("      CRICKET AUCTION & PDF-TO-PPTX CONVERTER TOOLKIT")
    print("=" * 64)


def handle_pdf_convert():
    print("\n--- [1] Convert PDF to PowerPoint Deck ---")
    pdf_path = ask_string("Source PDF path", DEFAULT_PDF_INPUT)
    ppt_path = ask_string("Output PPTX path", DEFAULT_PPTX_DIRECT_OUTPUT)
    aspect = ask_choice("Slide aspect ratio", ["match", "16:9", "4:3"], "match")
    fmt = ask_choice("Image format", ["jpeg", "png"], "jpeg")
    dpi = ask_int("DPI resolution (e.g. 100, 120, 150)", 120)
    pages = ask_string("Page ranges (Leave empty for all pages)", "")

    page_indices = None
    if pages:
        import pymupdf
        doc = pymupdf.open(pdf_path)
        page_indices = parse_page_spec(pages, len(doc))
        doc.close()

    run_conversion(pdf_path=pdf_path, output_pptx=ppt_path,
                   page_indices=page_indices, dpi=dpi,
                   img_format=fmt.lower(), aspect_ratio=aspect.lower())


def handle_page_select():
    print("\n--- [2] Select & Extract PDF Pages ---")
    pdf_path = ask_string("Source PDF path", DEFAULT_PDF_INPUT)
    output_pdf = ask_string("Output PDF path", DEFAULT_SELECTED_PDF_OUTPUT)
    page_spec = ask_string("Page specification (e.g., 1-4, 85-87, 182-185)", "1-4, 85-87, 182-185")

    run_page_selection(input_pdf=pdf_path, output_pdf=output_pdf, page_spec=page_spec)


def handle_pdf_compress():
    print("\n--- [3] Compress PDF File Size ---")
    pdf_path = ask_string("Source PDF path", DEFAULT_PDF_INPUT)
    output_pdf = ask_string("Output compressed PDF path", DEFAULT_COMPRESSED_PDF_OUTPUT)

    run_compression(input_pdf=pdf_path, output_pdf=output_pdf)


def handle_phase1_extract():
    print("\n--- [4] Phase 1: Extract PDF Player Stats & Photos ---")
    pdf_path = ask_string("Source PDF path", DEFAULT_PDF_INPUT)
    excel_path = ask_string("Output Excel path", DEFAULT_CURRENT_EXCEL)
    img_dir = ask_string("Directory to save player portraits", DEFAULT_CURRENT_IMG_DIR)

    run_pdf_extraction(pdf_path=pdf_path, output_excel=excel_path, img_dir=img_dir)


def handle_phase2_extract():
    print("\n--- [5] Phase 2: Extract Previous PPTX Auction Data ---")
    pptx_path = ask_string("Source PPTX path", DEFAULT_PPTX_INPUT)
    excel_path = ask_string("Output Excel path", DEFAULT_PREVIOUS_EXCEL)
    img_dir = ask_string("Directory to save images", DEFAULT_PREVIOUS_IMG_DIR)

    run_pptx_extraction(pptx_path=pptx_path, output_excel=excel_path, img_dir=img_dir)


def handle_phase3_merge():
    print("\n--- [6] Phase 3: Consolidate & Merge Datasets ---")
    curr_excel = ask_string("Current extracted Excel", DEFAULT_CURRENT_EXCEL)
    prev_excel = ask_string("Previous extracted Excel", DEFAULT_PREVIOUS_EXCEL)
    out_excel = ask_string("Master consolidated Excel", DEFAULT_MASTER_EXCEL)
    auto = ask_choice("Auto-merge matches without prompting?", ["Y", "N"], "Y") == "Y"

    run_merger(current_excel=curr_excel, previous_excel=prev_excel,
               output_excel=out_excel, auto_mode=auto)


def handle_phase4_generate():
    print("\n--- [7] Phase 4: Generate HPL 2.0 Auction Presentation ---")
    excel_path = ask_string("Master Excel path", DEFAULT_MASTER_EXCEL)
    output_pptx = ask_string("Output presentation path", "HPL_2_Auction_Deck.pptx")

    run_presentation_generation(excel_path=excel_path, output_pptx=output_pptx)


def handle_full_pipeline():
    print("\n" + "=" * 60)
    print("STARTING END-TO-END CRICKET AUCTION PIPELINE")
    print("=" * 60)
    print("\n>>> STEP 1: Extracting player stats & photos from PDF...")
    run_pdf_extraction(DEFAULT_PDF_INPUT, DEFAULT_CURRENT_EXCEL, DEFAULT_CURRENT_IMG_DIR)

    print("\n>>> STEP 2: Merging data into Master dataset...")
    run_merger(DEFAULT_CURRENT_EXCEL, DEFAULT_PREVIOUS_EXCEL, DEFAULT_MASTER_EXCEL, auto_mode=True)

    print("\n>>> STEP 3: Generating HPL 2.0 presentation deck...")
    run_presentation_generation(DEFAULT_MASTER_EXCEL, "HPL_2_Auction_Deck.pptx")

    print("\n[+] Full pipeline completed successfully!")


def handle_status():
    print("\n--- [9] Workspace File Status ---")
    tracked_files = [
        ("Source PDF", DEFAULT_PDF_INPUT),
        ("Previous PPTX", DEFAULT_PPTX_INPUT),
        ("Current Excel", DEFAULT_CURRENT_EXCEL),
        ("Previous Excel", DEFAULT_PREVIOUS_EXCEL),
        ("Master Excel", DEFAULT_MASTER_EXCEL),
        ("Direct PPTX Output", DEFAULT_PPTX_DIRECT_OUTPUT),
        ("HPL 2.0 Auction Deck", "HPL_2_Auction_Deck.pptx")
    ]
    for label, path in tracked_files:
        if os.path.exists(path):
            size_str = format_bytes(os.path.getsize(path))
            print(f"  [✓] {label:<22}: {os.path.basename(path)} ({size_str})")
        else:
            print(f"  [x] {label:<22}: Not Found ({os.path.basename(path)})")

    if os.path.exists(DEFAULT_CURRENT_IMG_DIR):
        img_count = len([f for f in os.listdir(DEFAULT_CURRENT_IMG_DIR) if f.lower().endswith(('.png', '.jpg'))])
        print(f"  [✓] Current Photos Dir    : {os.path.basename(DEFAULT_CURRENT_IMG_DIR)}/ ({img_count} photos)")


def launch_interactive_menu():
    """
    Main interactive loop.
    """
    actions = {
        "1": handle_pdf_convert,
        "2": handle_page_select,
        "3": handle_pdf_compress,
        "4": handle_phase1_extract,
        "5": handle_phase2_extract,
        "6": handle_phase3_merge,
        "7": handle_phase4_generate,
        "8": handle_full_pipeline,
        "9": handle_status
    }

    while True:
        print_banner()
        print("  [1] Convert PDF to PowerPoint (Direct Page Rasterizer)")
        print("  [2] Select / Extract Specific Pages from PDF")
        print("  [3] Compress PDF File Size")
        print("  [4] Phase 1: Extract Player Stats & Photos from PDF")
        print("  [5] Phase 2: Extract Previous PPTX Auction Data")
        print("  [6] Phase 3: Consolidate & Merge Datasets")
        print("  [7] Phase 4: Generate HPL 2.0 Auction Presentation (PPTX)")
        print("  [8] Run Complete Pipeline (Phase 1 -> 3 -> 4 End-to-End)")
        print("  [9] Inspect Workspace File Status")
        print("  [0] Exit")
        print("=" * 64)

        choice = input("Select an option [0-9]: ").strip()

        if choice == "0":
            print("\nExiting. Thank you!")
            sys.exit(0)
        elif choice in actions:
            try:
                actions[choice]()
            except Exception as e:
                print(f"\n[!] Error during execution: {e}")
            input("\nPress Enter to return to menu...")
        else:
            print("[!] Invalid option. Please select 0 to 9.")
