# Arizona 135 GT Frame Setup Tool

A standalone Tkinter desktop utility that automates the outline and artwork layout workflow for Canon Arizona 135 GT flatbed printing. The application replicates the SmartImprint workflow by generating two perfectly aligned PDFs:

- `<job>_ArizonaOutlines.pdf` – step-and-repeated outline clusters with configurable stroke and color.
- `<job>_ArizonaArtwork.pdf` – logo artwork positioned with bottom-anchored matte logic and optional flipping / rotation controls.

## Features

- Responsive ttk user interface with SRG-branded splash screen, live status bar, and progress feedback.
- Input controls for glass frame dimensions, frame indent, matte opening, bed width/height, horizontal and vertical cluster gaps, outline color, thickness, and output directory.
- Adjustable cluster columns and rows with live capacity calculations that flag when the layout exceeds the 96×48″ bed.
- Frame quantity selector that numbers the schedule order, allowing quick one-off exports or partial fills without re-entering measurements.
- EPS logo loader that preserves vector data via Ghostscript and reports detected logo size in millimetres.
- Options to flip artwork for reverse printing, rotate the bottom row 180°, and allow raster fallback when PyMuPDF cannot maintain vectors.
- Live dual preview showing the full bed arrangement (with inactive frames crosshatched) alongside a 1-up matte reference.
- Toggleable outline/artwork exports so operators can generate only the files required for the current press run.
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

1. Enter job parameters, choose outline colour, and adjust the cluster grid / frame quantity as required.
2. Load the EPS logo file (skip if exporting outlines only). The detected size is displayed in millimetres.
3. Pick output directory, configure artwork options, and choose which PDFs to export.
4. Click **Generate PDFs** to create the aligned outline and/or artwork files.

The generated PDFs are saved in the chosen directory using the job name slug with `ArizonaOutlines` and `ArizonaArtwork` suffixes.

## Notes

- Ghostscript is required for EPS to PDF conversion and must be installed separately.
- When raster fallback is enabled the tool produces high-resolution bitmaps only if PyMuPDF cannot embed the logo as vectors.
- No temporary files remain after export; Ghostscript conversions are streamed directly into memory.
