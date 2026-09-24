"""Small offline checks for PPTX integrity and PDF/PPTX matching."""

from __future__ import annotations

import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

from pypdf import PdfWriter


sys.path.insert(
    0,
    str(
        Path(__file__).resolve().parents[1]
        / "plugins" / "ppt-vector-export" / "skills" / "ppt-vector-export" / "scripts"
    ),
)
from export_figure import finish, prepare  # noqa: E402


PRESENTATION_XML = b"""<?xml version='1.0' encoding='UTF-8'?>
<p:presentation xmlns:p='http://schemas.openxmlformats.org/presentationml/2006/main'>
  <p:sldSz cx='914400' cy='914400'/>
</p:presentation>"""

SLIDE_XML = b"""<?xml version='1.0' encoding='UTF-8'?>
<p:sld xmlns:p='http://schemas.openxmlformats.org/presentationml/2006/main'
       xmlns:a='http://schemas.openxmlformats.org/drawingml/2006/main'>
  <p:cSld><p:spTree><p:sp><p:spPr>
    <a:xfrm><a:off x='10000' y='-5000'/>
      <a:ext cx='1270000' cy='635000'/></a:xfrm>
    <a:solidFill><a:srgbClr val='D9E7FA'/></a:solidFill>
  </p:spPr></p:sp></p:spTree></p:cSld>
</p:sld>"""


def make_pptx(path: Path) -> None:
    with zipfile.ZipFile(path, "w") as deck:
        deck.writestr("ppt/presentation.xml", PRESENTATION_XML)
        deck.writestr("ppt/slides/slide1.xml", SLIDE_XML)
        deck.writestr("ppt/media/original.png", b"original media bytes")


def make_pdf(path: Path, width: float, height: float) -> None:
    writer = PdfWriter()
    writer.add_blank_page(width=width, height=height)
    with path.open("wb") as stream:
        writer.write(stream)


class FigureExportTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.folder = Path(self.temp.name)
        self.source = self.folder / "source.pptx"
        self.canvas = self.folder / "canvas.pptx"
        make_pptx(self.source)

    def test_prepare_keeps_media_and_style_and_only_changes_geometry(self) -> None:
        result = prepare(self.source, self.canvas, padding_emu=12000)
        self.assertEqual(result["moved_objects"], 1)
        self.assertEqual(result["shift_emu"], [2000, 17000])
        self.assertEqual(result["slide_size_emu"], [1294000, 659000])
        with zipfile.ZipFile(self.source) as before, zipfile.ZipFile(self.canvas) as after:
            self.assertEqual(set(before.namelist()), set(after.namelist()))
            self.assertEqual(
                [name for name in before.namelist() if before.read(name) != after.read(name)],
                ["ppt/presentation.xml", "ppt/slides/slide1.xml"],
            )
            self.assertEqual(before.read("ppt/media/original.png"), after.read("ppt/media/original.png"))
            self.assertIn(b"D9E7FA", after.read("ppt/slides/slide1.xml"))

    def test_finish_rejects_pdf_from_a_different_slide(self) -> None:
        prepare(self.source, self.canvas, padding_emu=12000)
        native_pdf = self.folder / "wrong.pdf"
        make_pdf(native_pdf, 72, 72)
        output_pdf = self.folder / "result.pdf"
        output_svg = self.folder / "result.svg"
        with self.assertRaisesRegex(ValueError, "page size does not match"):
            finish(self.canvas, native_pdf, output_pdf, output_svg, None, None, False)
        self.assertFalse(output_pdf.exists())
        self.assertFalse(output_svg.exists())

    def test_finish_accepts_a_slide_without_embedded_svg(self) -> None:
        prepare(self.source, self.canvas, padding_emu=12000)
        native_pdf = self.folder / "native.pdf"
        make_pdf(native_pdf, 1294000 / 12700, 659000 / 12700)
        output_pdf = self.folder / "result.pdf"
        output_svg = self.folder / "result.svg"
        report = finish(self.canvas, native_pdf, output_pdf, output_svg, None, None, False)
        self.assertEqual(report["restored_svg_instances"], 0)
        self.assertTrue(output_pdf.is_file())
        self.assertTrue(output_svg.is_file())


if __name__ == "__main__":
    unittest.main()
