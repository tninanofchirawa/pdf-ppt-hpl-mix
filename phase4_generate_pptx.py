"""
Phase 4: HPL 2.0 Auction Presentation Generator Entrypoint
==========================================================
Builds the official widescreen (16:9) 'HPL 2.0' presentation deck.
"""

import sys
import os
import argparse

sys.path.insert(0, os.path.abspath(os.path.dirname(__file__)))

from generate_hpl2_deck import (
    generate_hpl2_presentation,
    DEFAULT_LEAGUE_TAG,
    DEFAULT_AUCTION_TITLE,
    DEFAULT_AUCTION_SUBTITLE,
    DEFAULT_EVENT_TAG,
    DEFAULT_ORGANIZER
)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Generate HPL 2.0 Mega Player Auction presentation")
    parser.add_argument("-i", "--input", default="outputs/Master_Cricket_Auction.xlsx", help="Master Excel data file")
    parser.add_argument("-o", "--output", default="HPL_2_Auction_Deck.pptx", help="Output PowerPoint presentation")
    parser.add_argument("--sample", action="store_true", help="Generate deck with marquee sample players")
    parser.add_argument("--title", default=DEFAULT_AUCTION_TITLE, help="League title")
    parser.add_argument("--subtitle", default=DEFAULT_AUCTION_SUBTITLE, help="League subtitle")
    parser.add_argument("--tagline", default=DEFAULT_LEAGUE_TAG, help="League tagline")
    parser.add_argument("--organizer", default=DEFAULT_ORGANIZER, help="Organizer footer")
    parser.add_argument("--event-tag", default=DEFAULT_EVENT_TAG, help="Event badge")
    args = parser.parse_args()

    data_source = "SAMPLE" if args.sample else args.input
    generate_hpl2_presentation(
        players_data=data_source,
        output_path=args.output,
        tagline=args.tagline,
        title=args.title,
        subtitle=args.subtitle,
        event_tag=args.event_tag,
        organizer=args.organizer
    )
