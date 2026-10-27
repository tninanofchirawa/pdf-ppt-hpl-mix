# System Architecture & Folder Organization

This project is structured around high cohesion and low coupling. All modules adhere to single-responsibility design principles, grouped into logical subfolders under parent purpose directories.

---

## Workspace Directory Layout

```text
pdf-ppt_convotor/
│
├── inputs/                          # Raw input source documents
│   ├── Cricket_Auction.pdf          # 253-page source auction PDF
│   └── IPL AUCTION.pptx             # Historical auction presentation deck
│
├── outputs/                         # All generated assets and datasets
│   ├── Cricket Auction.pptx         # Direct rasterized presentation deck
│   ├── Final_Cricket_Auction_2026.pptx # Branded, styled auction presentation deck
│   ├── Extracted_Current.xlsx       # Phase 1 current year structured dataset
│   ├── Extracted_Previous.xlsx      # Phase 2 historical structured dataset
│   ├── Master_Cricket_Auction.xlsx  # Phase 3 consolidated master dataset
│   ├── images_current/              # High-resolution extracted player portraits (Current)
│   └── images_previous/             # Extracted portrait candidates (Historical)
│
├── docs/                            # Documentation
│   ├── ARCHITECTURE.md              # System architecture & component design
│   ├── USER_GUIDE.md                # Interactive & CLI user workflows
│   └── MISTAKES_AND_FIXES.md        # Comprehensive bug autopsy and resolutions
│
├── src/                             # Core Python package
│   ├── common/                      # Shared configuration and utilities
│   │   ├── config/settings.py       # Paths, slide aspect ratios, theme palettes
│   │   └── utils/
│   │       ├── file_utils.py        # Multi-directory path resolver & formatters
│   │       └── text_utils.py        # Value sanitizer, string normalizer, fuzzy matcher
│   │
│   ├── pdf_tools/                   # General PDF operations
│   │   ├── converter/               # Direct PDF-to-PowerPoint rasterizer
│   │   │   ├── rasterizer.py        # PyMuPDF pixmap rendering with JPEG compression
│   │   │   ├── pptx_builder.py      # PowerPoint deck builder with dimension matching
│   │   │   └── runner.py            # Conversion orchestrator
│   │   ├── page_selector/           # Page extraction
│   │   │   ├── range_parser.py      # Bounds-checked 1-based range parser
│   │   │   ├── extractor.py         # Direct PDF page extraction
│   │   │   └── runner.py            # Page selection orchestrator
│   │   └── compressor/              # PDF compression
│   │       ├── optimizer.py         # Stream deflater & garbage collection
│   │       └── runner.py            # Compression metrics reporter
│   │
│   ├── auction_pipeline/            # Cricket Auction pipeline
│   │   ├── extractors/
│   │   │   ├── pdf/                 # Phase 1: PDF player extraction
│   │   │   │   ├── card_parser.py   # Text stats parser
│   │   │   │   ├── category_detector.py # Section and tag categorization
│   │   │   │   ├── portrait_extractor.py # Photo isolation & overseas detection
│   │   │   │   └── runner.py        # Phase 1 orchestrator
│   │   │   └── pptx/                # Phase 2: PPTX slide extraction
│   │   │       ├── shape_walker.py  # Recursive GroupShape traverser
│   │   │       ├── card_parser.py   # Text attribute parser
│   │   │       └── runner.py        # Phase 2 orchestrator
│   │   ├── merger/                  # Phase 3: Player deduplication & merge
│   │   │   ├── matcher.py           # Fuzzy player matcher & attribute enrichment
│   │   │   └── runner.py            # Phase 3 orchestrator
│   │   └── generator/               # Phase 4: Auction presentation generator
│   │       ├── styler.py            # Stadium Dark canvas setup & theme styling
│   │       ├── divider_slide.py     # Section title & countdown checkpoint slides
│   │       ├── player_card.py       # Player photo frame, name banner, metric chips
│   │       ├── stats_table.py       # Batting and bowling statistics tables
│   │       └── runner.py            # Phase 4 orchestrator
│   │
│   └── cli/                         # User Interface
│       ├── interactive/             # Interactive terminal menus
│       │   ├── prompts.py           # Smart prompt inputs with default fallbacks
│       │   └── menu.py              # Main terminal menu loop
│       └── commands/                # Command line subcommands
│           └── dispatcher.py        # Subcommand argument dispatcher
│
├── main.py                          # Primary entrypoint (Interactive or CLI)
├── run.py                           # Quick runner alias
└── requirements.txt                 # Project dependencies
```
