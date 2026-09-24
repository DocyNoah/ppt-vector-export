#!/usr/bin/env python3
"""Preserve a one-slide PowerPoint figure through native PDF export.

prepare: make an optional copy with a canvas fitted to its objects.
finish: restore SVG pictures rasterized by PowerPoint, then export SVG.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import posixpath
import shutil
import subprocess
import tempfile
import zipfile
from io import BytesIO
from pathlib import Path

from lxml import etree
from pypdf import PdfReader, PdfWriter
from pypdf.generic import (
    ArrayObject,
    ContentStream,
    DecodedStreamObject,
    DictionaryObject,
    NameObject,
    NumberObject,
)


EMU_PER_POINT = 12700
NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
}
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def run(*args: str, allowed: tuple[int, ...] = (0,)) -> subprocess.CompletedProcess[str]:
    result = subprocess.run(args, capture_output=True, text=True, check=False)
    if result.returncode not in allowed:
        raise RuntimeError(f"{' '.join(args)} failed ({result.returncode}):\n{result.stderr}")
    return result


def digest(path: Path) -> str:
    hasher = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            hasher.update(chunk)
    return hasher.hexdigest()


def require_new(*paths: Path) -> None:
    existing = [str(path) for path in paths if path.exists()]
    if existing:
        raise FileExistsError("Refusing to replace existing output: " + ", ".join(existing))


def slide_members(archive: zipfile.ZipFile) -> list[str]:
    names = archive.namelist()
    return sorted(
        name for name in names
        if name.startswith("ppt/slides/slide")
        and name.endswith(".xml")
        and "/_rels/" not in name
    )


def require_one_slide(archive: zipfile.ZipFile) -> None:
    slides = slide_members(archive)
    if slides != ["ppt/slides/slide1.xml"]:
        raise ValueError(f"Expected one slide named slide1.xml, found {slides}")


def transform_of(shape: etree._Element) -> etree._Element | None:
    for xpath in ("./p:spPr/a:xfrm", "./p:xfrm", "./p:grpSpPr/a:xfrm"):
        found = shape.xpath(xpath, namespaces=NS)
        if found:
            return found[0]
    return None


def slide_size_points(archive: zipfile.ZipFile) -> tuple[float, float]:
    presentation = etree.fromstring(archive.read("ppt/presentation.xml"))
    size = presentation.find("p:sldSz", NS)
    if size is None:
        raise ValueError("PPTX has no slide size")
    return int(size.get("cx")) / EMU_PER_POINT, int(size.get("cy")) / EMU_PER_POINT


def prepare(source: Path, output: Path, padding_emu: int) -> dict:
    require_new(output)
    if padding_emu < 0:
        raise ValueError("Padding must be nonnegative")
    with zipfile.ZipFile(source) as original:
        require_one_slide(original)
        slide = etree.fromstring(original.read("ppt/slides/slide1.xml"))
        presentation = etree.fromstring(original.read("ppt/presentation.xml"))
        tree = slide.xpath(".//p:spTree", namespaces=NS)[0]
        objects = []
        bounds = []
        for shape in tree:
            transform = transform_of(shape)
            if transform is None:
                continue
            offset = transform.find(f"{{{NS['a']}}}off")
            extent = transform.find(f"{{{NS['a']}}}ext")
            if offset is None or extent is None:
                continue
            x, y = int(offset.get("x")), int(offset.get("y"))
            width, height = int(extent.get("cx")), int(extent.get("cy"))
            objects.append((offset, x, y))
            bounds.append((x, y, x + width, y + height))
        if not bounds:
            raise ValueError("Slide has no positioned objects")
        min_x, min_y = min(b[0] for b in bounds), min(b[1] for b in bounds)
        max_x, max_y = max(b[2] for b in bounds), max(b[3] for b in bounds)
        dx, dy = padding_emu - min_x, padding_emu - min_y
        for offset, x, y in objects:
            offset.set("x", str(x + dx))
            offset.set("y", str(y + dy))
        slide_size = presentation.find("p:sldSz", NS)
        if slide_size is None:
            raise ValueError("PPTX has no slide size")
        slide_size.set("cx", str(max_x - min_x + 2 * padding_emu))
        slide_size.set("cy", str(max_y - min_y + 2 * padding_emu))
        changed = {
            "ppt/slides/slide1.xml": etree.tostring(
                slide, xml_declaration=True, encoding="UTF-8", standalone=True
            ),
            "ppt/presentation.xml": etree.tostring(
                presentation, xml_declaration=True, encoding="UTF-8", standalone=True
            ),
        }
        output.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(output, "w") as clone:
            for item in original.infolist():
                clone.writestr(item, changed.get(item.filename, original.read(item.filename)))
    with zipfile.ZipFile(output) as check:
        if check.testzip() is not None:
            raise RuntimeError("Generated PPTX failed ZIP integrity check")
    return {
        "source_sha256": digest(source),
        "export_pptx_sha256": digest(output),
        "moved_objects": len(objects),
        "shift_emu": [dx, dy],
        "slide_size_emu": [int(slide_size.get("cx")), int(slide_size.get("cy"))],
    }


def embedded_svg_pictures(pptx: Path) -> list[dict]:
    pictures = []
    with zipfile.ZipFile(pptx) as archive:
        require_one_slide(archive)
        slide = etree.fromstring(archive.read("ppt/slides/slide1.xml"))
        rels_name = "ppt/slides/_rels/slide1.xml.rels"
        if rels_name in archive.namelist():
            rels = etree.fromstring(archive.read(rels_name))
            targets = {item.get("Id"): item.get("Target") for item in rels}
        else:
            targets = {}
        for picture in slide.xpath(".//p:pic", namespaces=NS):
            svg_blips = picture.xpath('.//*[local-name()="svgBlip"]')
            if not svg_blips:
                continue
            rid = svg_blips[0].get(f"{{{REL}}}embed")
            if rid not in targets:
                raise ValueError(f"Missing relationship for embedded SVG {rid}")
            target = targets[rid]
            name = (
                target.lstrip("/") if target.startswith("/")
                else posixpath.normpath(posixpath.join("ppt/slides", target))
            )
            if not name.lower().endswith(".svg"):
                raise ValueError(f"SVG relationship points to {name}")
            transform = picture.find("p:spPr/a:xfrm", NS)
            if transform is None:
                raise ValueError(f"Embedded SVG {name} has no position transform")
            offset = transform.find("a:off", NS)
            extent = transform.find("a:ext", NS)
            if offset is None or extent is None:
                raise ValueError(f"Embedded SVG {name} has incomplete coordinates")
            pictures.append({
                "name": name,
                "x": int(offset.get("x")) / EMU_PER_POINT,
                "y": int(offset.get("y")) / EMU_PER_POINT,
                "width": int(extent.get("cx")) / EMU_PER_POINT,
                "height": int(extent.get("cy")) / EMU_PER_POINT,
                "svg": archive.read(name),
            })
    return pictures


def restore_svg_images(pptx: Path, native_pdf: Path, final_pdf: Path) -> int:
    import cairosvg

    pictures = embedded_svg_pictures(pptx)
    if not pictures:
        shutil.copyfile(native_pdf, final_pdf)
        return 0
    reader = PdfReader(native_pdf)
    writer = PdfWriter()
    page = writer.add_page(reader.pages[0])
    xobjects = page["/Resources"].get("/XObject")
    if xobjects is None:
        raise ValueError("PDF has no image XObjects to replace")
    xobjects = xobjects.get_object()
    content = ContentStream(page.get_contents(), writer)
    page_height = float(page.mediabox.height)
    candidates = []
    for index, (operands, operator) in enumerate(content.operations):
        if operator != b"Do" or index == 0 or content.operations[index - 1][1] != b"cm":
            continue
        name = operands[0]
        if name not in xobjects or xobjects[name].get_object().get("/Subtype") != "/Image":
            continue
        matrix = content.operations[index - 1][0]
        if len(matrix) != 6:
            continue
        width, height, x, bottom = (float(matrix[i]) for i in (0, 3, 4, 5))
        candidates.append((index, name, x, page_height - bottom - height, width, height))

    matched = set()
    replaced_names = set()
    for number, picture in enumerate(pictures, start=1):
        nearby = sorted(
            (abs(x - picture["x"]) + abs(y - picture["y"]), index)
            for index, _, x, y, width, height in candidates
            if index not in matched
            and abs(width - picture["width"]) <= 1
            and abs(height - picture["height"]) <= 1
        )
        if not nearby or nearby[0][0] > 0.75:
            raise RuntimeError(
                f"Cannot match embedded SVG {picture['name']} at "
                f"({picture['x']:.2f}, {picture['y']:.2f}) pt"
            )
        if len(nearby) > 1 and nearby[1][0] - nearby[0][0] < 0.25:
            raise RuntimeError(f"Ambiguous PDF image match for {picture['name']}")
        content_index = nearby[0][1]
        matched.add(content_index)
        original_svg = etree.fromstring(picture["svg"])
        original_svg.set("width", "1pt")
        original_svg.set("height", "1pt")
        original_svg.set("preserveAspectRatio", "none")
        vector_pdf = cairosvg.svg2pdf(bytestring=etree.tostring(original_svg))
        vector_page = PdfReader(BytesIO(vector_pdf)).pages[0]
        form = DecodedStreamObject()
        form.set_data(vector_page.get_contents().get_data())
        form.update(DictionaryObject({
            NameObject("/Type"): NameObject("/XObject"),
            NameObject("/Subtype"): NameObject("/Form"),
            NameObject("/FormType"): NumberObject(1),
            NameObject("/BBox"): ArrayObject([
                NumberObject(0), NumberObject(0), NumberObject(1), NumberObject(1)
            ]),
            NameObject("/Resources"): vector_page["/Resources"].clone(writer),
        }))
        new_name = NameObject(f"/RestoredSVG{number}")
        if new_name in xobjects:
            raise RuntimeError(f"PDF already uses XObject name {new_name}")
        xobjects[new_name] = writer._add_object(form)
        old_name = content.operations[content_index][0][0]
        content.operations[content_index][0][0] = new_name
        replaced_names.add(old_name)

    used_names = {
        operands[0] for operands, operator in content.operations
        if operator == b"Do" and operands
    }
    for name in replaced_names - used_names:
        del xobjects[name]
    page[NameObject("/Contents")] = writer._add_object(content)
    with final_pdf.open("wb") as stream:
        writer.write(stream)
    return len(matched)


def image_object_count(pdf: Path) -> int:
    lines = run("pdfimages", "-list", str(pdf)).stdout.splitlines()
    return max(0, len(lines) - 2)


def compare_png(reference: Path, pdf: Path, svg: Path, qa_dir: Path) -> dict:
    import cairosvg
    from PIL import Image, ImageChops, ImageStat

    def on_white(image: Image.Image) -> Image.Image:
        rgba = image.convert("RGBA")
        white = Image.new("RGBA", rgba.size, "white")
        return Image.alpha_composite(white, rgba).convert("RGB")

    qa_dir.mkdir(parents=True, exist_ok=True)
    with Image.open(reference) as raw:
        source = on_white(raw)
    width, height = source.size
    prefix = qa_dir / "pdf-render"
    run(
        "pdftoppm", "-f", "1", "-singlefile", "-png",
        "-scale-to-x", str(width), "-scale-to-y", str(height),
        str(pdf), str(prefix),
    )
    with Image.open(prefix.with_suffix(".png")) as raw:
        rendered = on_white(raw)
    if rendered.size != source.size:
        raise ValueError(f"Rendered size {rendered.size} differs from reference {source.size}")
    difference = ImageChops.difference(source, rendered)
    difference.save(qa_dir / "difference.png")
    Image.blend(source, rendered, 0.5).save(qa_dir / "overlay.png")
    prefix.with_suffix(".png").unlink()
    means = ImageStat.Stat(difference).mean
    svg_render = qa_dir / "svg-render.png"
    cairosvg.svg2png(
        url=str(svg), write_to=str(svg_render),
        output_width=width, output_height=height,
    )
    with Image.open(svg_render) as raw:
        svg_pixels = on_white(raw)
    if svg_pixels.size != source.size:
        raise ValueError(f"SVG rendered size {svg_pixels.size} differs from reference {source.size}")
    svg_pdf_means = ImageStat.Stat(ImageChops.difference(rendered, svg_pixels)).mean
    return {
        "reference_size": [width, height],
        "reference_sha256": digest(reference),
        "rgb_mean_absolute_error": round(sum(means) / 3, 4),
        "pdf_svg_rgb_mean_absolute_error": round(sum(svg_pdf_means) / 3, 4),
    }


def finish(
    pptx: Path, native_pdf: Path, output_pdf: Path, output_svg: Path,
    reference_png: Path | None, qa_dir: Path | None, skip_svg_restoration: bool,
) -> dict:
    require_new(output_pdf, output_svg)
    if qa_dir is not None and qa_dir.exists():
        raise FileExistsError(f"QA directory already exists: {qa_dir}")
    if qa_dir is not None and reference_png is None:
        raise ValueError("--qa-dir requires --reference-png")
    with zipfile.ZipFile(pptx) as archive:
        require_one_slide(archive)
        expected_width, expected_height = slide_size_points(archive)
    output_pdf.parent.mkdir(parents=True, exist_ok=True)
    output_svg.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="ppt-vector-export-") as folder:
        scratch = Path(folder)
        cleaned = scratch / "native-clean.pdf"
        repaired = run("qpdf", str(native_pdf), str(cleaned), allowed=(0, 3))
        if repaired.returncode == 3 and not cleaned.exists():
            raise RuntimeError("qpdf reported a warning but produced no repaired PDF")
        run("qpdf", "--check", str(cleaned))
        reader = PdfReader(cleaned)
        if len(reader.pages) != 1:
            raise ValueError(f"Expected one PDF page, found {len(reader.pages)}")
        page = reader.pages[0]
        actual_width = float(page.mediabox.width)
        actual_height = float(page.mediabox.height)
        if max(abs(actual_width - expected_width), abs(actual_height - expected_height)) > 1:
            raise ValueError(
                "PDF page size does not match PPTX slide: "
                f"PDF {actual_width:.2f}x{actual_height:.2f} pt, "
                f"PPTX {expected_width:.2f}x{expected_height:.2f} pt"
            )
        final_pdf = scratch / "figure.pdf"
        if skip_svg_restoration:
            shutil.copyfile(cleaned, final_pdf)
            restored = 0
        else:
            restored = restore_svg_images(pptx, cleaned, final_pdf)
        run("qpdf", "--check", str(final_pdf))
        final_svg = scratch / "figure.svg"
        run("pdftocairo", "-svg", str(final_pdf), str(final_svg))
        svg_root = etree.parse(str(final_svg)).getroot()
        svg_images = len(svg_root.xpath('.//*[local-name()="image"]'))
        svg_paths = len(svg_root.xpath('.//*[local-name()="path"]'))
        report = {
            "source_pptx_sha256": digest(pptx),
            "native_pdf_sha256": digest(native_pdf),
            "final_pdf_sha256": digest(final_pdf),
            "final_svg_sha256": digest(final_svg),
            "restored_svg_instances": restored,
            "svg_restoration_skipped": skip_svg_restoration,
            "remaining_pdf_image_objects_including_masks": image_object_count(final_pdf),
            "svg_path_elements": svg_paths,
            "svg_image_elements": svg_images,
            "page_size_pt": [actual_width, actual_height],
        }
        if reference_png is not None:
            report.update(compare_png(reference_png, final_pdf, final_svg, qa_dir or scratch / "qa"))
        shutil.copy2(final_pdf, output_pdf)
        shutil.copy2(final_svg, output_svg)
    if qa_dir is not None:
        (qa_dir / "report.json").write_text(json.dumps(report, indent=2) + "\n")
    return report


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest="command", required=True)
    before = commands.add_parser("prepare", help="Fit one PPTX slide to its objects")
    before.add_argument("--source", required=True, type=Path)
    before.add_argument("--output", required=True, type=Path)
    before.add_argument("--padding-emu", type=int, default=12000)
    after = commands.add_parser("finish", help="Restore embedded SVGs and create PDF/SVG")
    after.add_argument("--pptx", required=True, type=Path)
    after.add_argument("--native-pdf", required=True, type=Path)
    after.add_argument("--output-pdf", required=True, type=Path)
    after.add_argument("--output-svg", required=True, type=Path)
    after.add_argument("--reference-png", type=Path)
    after.add_argument("--qa-dir", type=Path)
    after.add_argument(
        "--skip-svg-restoration", action="store_true",
        help="Use only after confirming the native PDF already preserves SVG artwork",
    )
    args = parser.parse_args()
    if args.command == "prepare":
        result = prepare(args.source, args.output, args.padding_emu)
    else:
        result = finish(
            args.pptx, args.native_pdf, args.output_pdf, args.output_svg,
            args.reference_png, args.qa_dir, args.skip_svg_restoration,
        )
    print(json.dumps(result, indent=2))


if __name__ == "__main__":
    main()
