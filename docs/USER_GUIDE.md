# User Guide & Operations Manual

This guide covers all operating modes for the Cricket Auction & PDF-to-PPTX Converter Toolkit.

---

## 1. Interactive Menu Mode

The easiest way to run any part of the project is through the interactive menu:

```powershell
python main.py
# or
python run.py
```

```text
================================================================
      CRICKET AUCTION & PDF-TO-PPTX CONVERTER TOOLKIT
================================================================
  [1] Convert PDF to PowerPoint (Direct Page Rasterizer)
  [2] Select / Extract Specific Pages from PDF
  [3] Compress PDF File Size
  [4] Phase 1: Extract Player Stats & Photos from PDF
  [5] Phase 2: Extract Previous PPTX Auction Data
  [6] Phase 3: Consolidate & Merge Datasets
  [7] Phase 4: Generate Branded Auction Presentation (PPTX)
  [8] Run Complete Pipeline (Phase 1 -> 3 -> 4 End-to-End)
  [9] Inspect Workspace File Status
  [0] Exit
================================================================
```

### Navigating Options:
- Type a number from `0` to `9` and press **Enter**.
- Prompts show smart defaults inside brackets (e.g., `[Default: inputs\Cricket_Auction.pdf]`). Press **Enter** to accept the default or type a custom value.
- Option `[9]` inspects your `inputs/` and `outputs/` folders, displaying file sizes and photo counts.

---

## 2. Command Line Interface (CLI)

You can run individual actions without opening the interactive menu:

### Direct PDF-to-PowerPoint Conversion (`convert`)
Converts PDF pages directly to PowerPoint slides:
```powershell
# Convert entire PDF preserving original 4:3 dimensions
python main.py convert

# Convert specific pages (e.g., pages 1 to 10)
python main.py convert --pages "1-10" -o "outputs/Sample.pptx"

# Convert with custom DPI and 16:9 widescreen format
python main.py convert --aspect-ratio 16:9 --dpi 150
```

### Extracting Specific Pages (`select`)
Extracts specific pages to a new PDF:
```powershell
python main.py select -p "1-4, 85-87, 182-185" -o "outputs/Cricket_Auction_Selected.pdf"
```

### Compressing PDF (`compress`)
Compresses and optimizes PDF size:
```powershell
python main.py compress -i "inputs/Cricket_Auction.pdf" -o "outputs/compressed.pdf"
```

### Full End-to-End Auction Pipeline (`pipeline`)
Runs extraction, data consolidation, and presentation generation in a single automated step:
```powershell
python main.py pipeline
```

### Running Individual Pipeline Phases
```powershell
# Phase 1: PDF player extraction
python main.py extract-pdf

# Phase 2: PPTX slide extraction
python main.py extract-pptx

# Phase 3: Dataset merge (auto mode)
python main.py merge --auto

# Phase 4: Presentation deck generation
python main.py generate-pptx
```
