---
name: ppt-vector-export
description: Use when exporting a one-slide PPTX diagram or paper figure to vector PDF and SVG, especially if PowerPoint clips off-slide objects or rasterizes embedded SVG artwork.
---

# PowerPoint figure to vector PDF and SVG

Use Microsoft PowerPoint's **local** PDF export as the source of the drawing. Converting that PDF to SVG preserves its vector paths, but it cannot turn rasterized content back into vectors. `scripts/export_figure.py` can restore SVG pictures embedded in a one-slide PPTX when PowerPoint rasterizes them in the PDF.

## Workflow

1. Inspect the source PPTX and choose the output canvas. If any objects lie outside the slide or the figure needs a tight crop, run `prepare` to make a copy with a canvas enclosing all objects. It moves every top-level object by the same offset and changes the slide size; it does not edit the source or object styles. Otherwise export the source PPTX directly.
   The default padding is small; if line ends or arrowheads touch the export edge, rerun `prepare` with a larger `--padding-emu` and export that copy again.
2. Open the chosen PPTX in **Microsoft PowerPoint** and export a PDF using **Best for printing / 인쇄에 적합**. Use the local exporter. A different renderer, including LibreOffice, may change arrowheads, fonts, or other drawing details. Keep this native PDF as a checkpoint.
3. Run `finish` with the exact PPTX used in step 2 and the native PDF. It checks the page size, restores any embedded SVG pictures whose raster replacements match the PPTX positions, writes the final PDF, converts that PDF to SVG, and validates both. If matching fails, inspect the source and PDF; do not accept a partial replacement.
   If the native PDF already keeps every embedded SVG as vectors, pass `--skip-svg-restoration` after inspecting the PDF. This preserves the native PDF without an unnecessary replacement attempt.
4. Render the final PDF and SVG at the same dimensions as a trusted reference PNG or screenshot. Inspect arrows, text, colors, formula strokes, clipping, and any remaining raster images. Use `--reference-png` to save an overlay and difference image when a reference exists. Pixel differences from antialiasing alone are expected; visible shape changes are not.
5. Deliver the final PDF and SVG with the native PDF and the export PPTX when provenance matters. Keep the original PPTX untouched. Do not replace paper or website assets, commit, or publish unless the user requested that separately.

## Commands

Requires macOS PowerPoint, Python packages in `requirements.txt`, and `qpdf`, `pdftocairo`, `pdftoppm`, and `pdfimages` on `PATH`. CairoSVG also needs a working Cairo library. Resolve script and dependency paths relative to this `SKILL.md`; run the examples from its directory with a project-local virtual environment if dependencies are absent.

```sh
python scripts/export_figure.py prepare \
  --source '/absolute/path/source.pptx' \
  --output '/absolute/path/export-canvas.pptx'

# Open export-canvas.pptx in PowerPoint and export native.pdf locally.
python scripts/export_figure.py finish \
  --pptx '/absolute/path/export-canvas.pptx' \
  --native-pdf '/absolute/path/native.pdf' \
  --output-pdf '/absolute/path/figure-vector.pdf' \
  --output-svg '/absolute/path/figure-vector.svg' \
  --reference-png '/absolute/path/reference.png' \
  --qa-dir '/absolute/path/qa'
```

Omit `--reference-png` and `--qa-dir` if no reference exists. If the slide already has the right canvas, skip `prepare` and pass the original PPTX to `finish`. The script supports **one slide**; for a multi-slide deck, export with PowerPoint and convert pages individually only after inspecting each page's vector content.

## Fidelity limits

- A PDF or SVG wrapper does not make embedded photos or other original bitmaps vector. `finish` reports how many SVG instances were restored and how many image objects remain.
- SVG restoration relies on a recognizable PowerPoint image drawing operation at the expected position. The script stops on missing or ambiguous matches. Review the native PDF and adapt the method for different PowerPoint output structures.
- Check color in the final destination PDF or webpage, not only in the standalone figure. PNG inclusion can discard PNG color metadata, while a vector PDF may retain an ICC profile.
