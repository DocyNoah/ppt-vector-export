# PowerPoint Vector Export

A Codex plugin and Agent Skill for exporting one-slide PowerPoint figures to vector-preserving PDF and SVG. It keeps PowerPoint's native shape, text, and arrow rendering, then restores SVG pictures that PowerPoint rasterizes during PDF export.

## Install in Codex

Add this Git repository as a plugin marketplace and install the plugin:

```sh
codex plugin marketplace add DocyNoah/ppt-vector-export
codex plugin add ppt-vector-export@docynoah
```

Start a new Codex task after installation so the skill is available there. Ask Codex to export a one-slide PPTX figure as PDF and SVG, or invoke the `ppt-vector-export` skill directly. The skill instructions are in [`SKILL.md`](plugins/ppt-vector-export/skills/ppt-vector-export/SKILL.md).

This repository contains a [Codex marketplace catalog](.agents/plugins/marketplace.json), a [plugin manifest](plugins/ppt-vector-export/.codex-plugin/plugin.json), and the Agent Skill packaged inside the plugin. The `SKILL.md` and its script can also be used directly by agents that support the Agent Skills file format.

## What the workflow does

1. Optionally make a copy of the PPTX with a slide canvas fitted to all diagram objects.
2. Export that PPTX to PDF with **Microsoft PowerPoint's local “Best for printing” exporter**. This remains a manual step to preserve PowerPoint's own rendering of shapes and arrowheads.
3. Match any SVG pictures rasterized in the PDF to their positions in the PPTX, replace them with their original vector artwork, and derive an SVG from the resulting PDF.
4. Validate the PDF and SVG. With a reference PNG, generate an overlay, difference image, and comparison report.

The tool supports **one-slide PPTX figures**. Original bitmap photos or depth images remain bitmaps inside the final PDF/SVG. The command fails if an embedded SVG cannot be matched reliably; it does not silently produce a partially restored figure.

## Command-line use

Requirements: macOS PowerPoint, Python 3.12+, Cairo, Poppler, and qpdf. A Homebrew setup is:

```sh
brew install python@3.13 cairo poppler qpdf
python3.13 -m venv .venv
SKILL_DIR=plugins/ppt-vector-export/skills/ppt-vector-export
.venv/bin/python -m pip install -r "$SKILL_DIR/requirements.txt"
```

On Apple Silicon, CairoSVG may require `DYLD_LIBRARY_PATH=/opt/homebrew/lib`. Use the corresponding Homebrew library path on Intel Macs.

When the slide needs to include off-slide objects, create an export copy:

```sh
.venv/bin/python "$SKILL_DIR/scripts/export_figure.py" prepare \
  --source '/path/to/source.pptx' \
  --output '/path/to/export-canvas.pptx'
```

Open the PPTX you will export in PowerPoint. Choose **File → Export → PDF → Best for printing** (파일 → 내보내기 → PDF → 인쇄에 적합) and save the result as `native.pdf`. Use that same PPTX in the next command:

```sh
.venv/bin/python "$SKILL_DIR/scripts/export_figure.py" finish \
  --pptx '/path/to/export-canvas.pptx' \
  --native-pdf '/path/to/native.pdf' \
  --output-pdf '/path/to/figure-vector.pdf' \
  --output-svg '/path/to/figure-vector.svg' \
  --reference-png '/path/to/reference.png' \
  --qa-dir '/path/to/qa'
```

The reference PNG and QA directory are optional. If the source PPTX already has the right canvas, skip `prepare` and pass the source PPTX to `finish`. If PowerPoint's PDF already preserves all embedded SVGs as vectors, inspect it and use `--skip-svg-restoration`.

Inspect the final PDF and SVG at the intended display size. Compare arrowheads, equations, text, colors, and clipping with the source. Pixel differences from antialiasing can occur even when geometry is correct.

## Repository structure

```text
.agents/plugins/marketplace.json                Codex marketplace catalog
plugins/ppt-vector-export/
  .codex-plugin/plugin.json                     Codex plugin manifest
  skills/ppt-vector-export/
    SKILL.md                                     Agent Skill instructions
    agents/openai.yaml                          Skill UI metadata
    requirements.txt                            Python dependencies
    scripts/export_figure.py                    Prepare and finish commands
tests/test_export_figure.py                     Offline behavior checks
```

Run the offline checks with `python -m unittest discover -s tests -v` after installing the Python dependencies and system tools.
