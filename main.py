"""
Main Application Entrypoint
===========================
Interactive Launcher & CLI Interface for the Cricket Auction & PDF/PPTX Toolkit.

Usage:
  Interactive Mode : python main.py
  CLI Mode         : python main.py <command> [options]
"""

import sys
import os

# Ensure project root is in sys.path
sys.path.insert(0, os.path.abspath(os.path.dirname(__file__)))

from src.cli.commands.dispatcher import build_parser, dispatch_cli_command
from src.cli.interactive.menu import launch_interactive_menu


def main():
    if len(sys.argv) > 1:
        parser = build_parser()
        args = parser.parse_args()
        if args.subcommand:
            dispatch_cli_command(args)
        else:
            parser.print_help()
    else:
        # Launch interactive menu when invoked without arguments
        launch_interactive_menu()


if __name__ == "__main__":
    main()
