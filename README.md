# Arizona 135 GT Frame Setup Tool

A standalone Tkinter desktop utility that automates the outline and artwork layout workflow for Canon Arizona 135 GT flatbed printing. The application replicates the SmartImprint workflow by generating two perfectly aligned PDFs:

- `<job>_ArizonaOutlines.pdf` – step-and-repeated outline clusters with configurable stroke and colour.
- `<job>_ArizonaArtwork.pdf` – logo artwork positioned with matte-band logic and optional flipping / rotation controls.

## Features

- SRG-branded splash screen followed by a responsive ttk user interface with status bar and export progress feedback.
- Input controls for glass frame dimensions, frame indent, matte opening, bed width, cluster gap, outline colour / thickness, and output directory.
- Spinbox selector for cluster count with live capacity calculation based on the current bed width.
- EPS logo loader that converts through Ghostscript, keeps vector fidelity, and reports the detected logo size in millimetres.
- Options to flip artwork for reverse printing, rotate the bottom row 180°, and allow raster fallback when PyMuPDF cannot maintain vectors.
- Live preview that shows the bed width, clusters, matte band, logo placement, and rotation cues as parameters change.
- Outline and artwork exports share identical page dimensions and origin for guaranteed alignment on press.

## Requirements

- Python 3.10+
- Tkinter (bundled with standard CPython builds)
- [Ghostscript](https://www.ghostscript.com/) installed and available on the system `PATH`
- Python dependencies listed in `requirements.txt`

Install Python dependencies with:

```bash
pip install -r requirements.txt
```

## Running the tool

```bash
python main.py
```

1. Enter job parameters, choose outline colour, and adjust the cluster count if desired.
2. Load the EPS logo file. The detected size is displayed in millimetres.
3. Pick an output directory and configure the artwork options.
4. Click **Generate PDFs** to create the aligned outline and artwork files.

The generated PDFs are saved in the chosen directory using the job name slug with `ArizonaOutlines` and `ArizonaArtwork` suffixes.

## Notes

- Ghostscript is required for EPS to PDF conversion and must be installed separately.
- When raster fallback is enabled the tool produces a high-resolution bitmap only if PyMuPDF cannot embed the logo as vectors.
- No temporary files remain after export; Ghostscript conversions are streamed directly into memory.
