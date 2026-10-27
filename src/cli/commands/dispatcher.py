"""
Command Line Dispatcher
=======================
Parses CLI subcommands and executes the requested action without prompting.
"""

import argparse

from src.common.config.settings import (
    DEFAULT_PDF_INPUT, DEFAULT_PPTX_INPUT, DEFAULT_PPTX_DIRECT_OUTPUT,
    DEFAULT_SELECTED_PDF_OUTPUT, DEFAULT_COMPRESSED_PDF_OUTPUT,
    DEFAULT_CURRENT_EXCEL, DEFAULT_PREVIOUS_EXCEL, DEFAULT_MASTER_EXCEL,
    DEFAULT_FINAL_PPTX, DEFAULT_CURRENT_IMG_DIR, DEFAULT_PREVIOUS_IMG_DIR
)
from src.pdf_tools.converter import run_conversion
from src.pdf_tools.page_selector import run_page_selection, parse_page_spec
from src.pdf_tools.compressor import run_compression
from src.auction_pipeline.extractors.pdf import run_pdf_extraction
from src.auction_pipeline.extractors.pptx import run_pptx_extraction
from src.auction_pipeline.merger import run_merger
from src.auction_pipeline.generator import run_presentation_generation


def build_parser() -> argparse.ArgumentParser:
    """
    Constructs the command-line argument parser with all subcommands.
    """
    parser = argparse.ArgumentParser(
        description="Cricket Auction & PDF-to-PPTX Converter Toolkit",
        formatter_class=argparse.RawTextHelpFormatter
    )
    subparsers = parser.add_subparsers(dest="subcommand", help="Available subcommands")

    # convert
    p_conv = subparsers.add_parser("convert", help="Convert PDF pages to PowerPoint slides")
    p_conv.add_argument("-i", "--input", default=DEFAULT_PDF_INPUT, help="Source PDF")
    p_conv.add_argument("-o", "--output", default=DEFAULT_PPTX_DIRECT_OUTPUT, help="Target PPTX")
    p_conv.add_argument("--dpi", type=int, default=120, help="Rendering DPI (default: 120)")
    p_conv.add_argument("--format", choices=["jpeg", "png"], default="jpeg", help="Image format (default: jpeg)")
    p_conv.add_argument("--aspect-ratio", choices=["match", "16:9", "4:3"], default="match",
                        help="Slide aspect ratio mode (default: match exact PDF page dimensions)")
    p_conv.add_argument("--pages", default=None, help="Page range (e.g. '1-10, 25')")

    # select
    p_sel = subparsers.add_parser("select", help="Extract specific page ranges to a new PDF")
    p_sel.add_argument("-i", "--input", default=DEFAULT_PDF_INPUT, help="Source PDF")
    p_sel.add_argument("-o", "--output", default=DEFAULT_SELECTED_PDF_OUTPUT, help="Target PDF")
    p_sel.add_argument("-p", "--pages", default="1-4, 85-87, 182-185", help="Page range specification")

    # compress
    p_comp = subparsers.add_parser("compress", help="Compress PDF size")
    p_comp.add_argument("-i", "--input", default=DEFAULT_PDF_INPUT, help="Source PDF")
    p_comp.add_argument("-o", "--output", default=DEFAULT_COMPRESSED_PDF_OUTPUT, help="Target PDF")

    # extract-pdf
    p_epdf = subparsers.add_parser("extract-pdf", help="Phase 1: Extract PDF player stats & photos")
    p_epdf.add_argument("--pdf", default=DEFAULT_PDF_INPUT, help="Input PDF file")
    p_epdf.add_argument("--output", default=DEFAULT_CURRENT_EXCEL, help="Output Excel file")
    p_epdf.add_argument("--img-dir", default=DEFAULT_CURRENT_IMG_DIR, help="Directory for photos")

    # extract-pptx
    p_epptx = subparsers.add_parser("extract-pptx", help="Phase 2: Extract PPTX data")
    p_epptx.add_argument("--pptx", default=DEFAULT_PPTX_INPUT, help="Input PPTX file")
    p_epptx.add_argument("--output", default=DEFAULT_PREVIOUS_EXCEL, help="Output Excel file")
    p_epptx.add_argument("--img-dir", default=DEFAULT_PREVIOUS_IMG_DIR, help="Directory for images")

    # merge
    p_merge = subparsers.add_parser("merge", help="Phase 3: Merge and deduplicate player data")
    p_merge.add_argument("--current", default=DEFAULT_CURRENT_EXCEL, help="Current Excel")
    p_merge.add_argument("--previous", default=DEFAULT_PREVIOUS_EXCEL, help="Previous Excel")
    p_merge.add_argument("--output", default=DEFAULT_MASTER_EXCEL, help="Output Master Excel")
    p_merge.add_argument("--auto", action="store_true", help="Auto-merge without prompts")

    # generate-pptx
    p_gen = subparsers.add_parser("generate-pptx", help="Phase 4: Generate auction presentation")
    p_gen.add_argument("--excel", default=DEFAULT_MASTER_EXCEL, help="Master Excel")
    p_gen.add_argument("--output", default=DEFAULT_FINAL_PPTX, help="Output PPTX")

    # pipeline
    subparsers.add_parser("pipeline", help="Run full Phase 1 -> 3 -> 4 pipeline end-to-end")

    return parser


def dispatch_cli_command(args: argparse.Namespace) -> None:
    """
    Executes the action corresponding to the parsed subcommand.
    """
    if args.subcommand == "convert":
        page_indices = None
        if args.pages:
            import pymupdf
            doc = pymupdf.open(args.input)
            page_indices = parse_page_spec(args.pages, len(doc))
            doc.close()

        run_conversion(pdf_path=args.input, output_pptx=args.output,
                       page_indices=page_indices, dpi=args.dpi,
                       img_format=args.format, aspect_ratio=args.aspect_ratio)

    elif args.subcommand == "select":
        run_page_selection(input_pdf=args.input, output_pdf=args.output, page_spec=args.pages)

    elif args.subcommand == "compress":
        run_compression(input_pdf=args.input, output_pdf=args.output)

    elif args.subcommand == "extract-pdf":
        run_pdf_extraction(pdf_path=args.pdf, output_excel=args.output, img_dir=args.img_dir)

    elif args.subcommand == "extract-pptx":
        run_pptx_extraction(pptx_path=args.pptx, output_excel=args.output, img_dir=args.img_dir)

    elif args.subcommand == "merge":
        run_merger(current_excel=args.current, previous_excel=args.previous,
                   output_excel=args.output, auto_mode=args.auto)

    elif args.subcommand == "generate-pptx":
        run_presentation_generation(excel_path=args.excel, output_pptx=args.output)

    elif args.subcommand == "pipeline":
        from src.cli.interactive.menu import handle_full_pipeline
        handle_full_pipeline()
