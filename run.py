"""
Quick Runner Alias
==================
Delegates directly to main.py. Run with or without arguments:
  python run.py
  python run.py convert --pages "1-5"
"""

import sys
import os

# Ensure project root is in sys.path
sys.path.insert(0, os.path.abspath(os.path.dirname(__file__)))

from main import main

if __name__ == "__main__":
    main()
