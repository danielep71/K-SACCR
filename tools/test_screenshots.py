"""Tests of tools/screenshots.py, the processing of the documentation screenshots (#616)."""
from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

import screenshots

try:
    from PIL import Image  # type: ignore[import-not-found, unused-ignore]
except ImportError:  # pragma: no cover - Pillow is optional for the other tools
    Image = None


@unittest.skipIf(Image is None, "Pillow is not installed")
class ScreenshotTests(unittest.TestCase):
    def test_only_the_documentation_pictures_are_written_as_a_palette(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            source = Path(folder) / "screenshots"
            source.mkdir()
            picture = Image.new("RGB", (40, 20), (255, 255, 255))
            for x in range(40):
                picture.putpixel((x, 5), (x * 6, 56, 100))
            picture.save(source / "SACCR_Results_EAD.png")
            picture.save(source / "busy-dialog-1.png")
            root = Path(folder) / "repo"
            self.assertEqual([p.name for p in screenshots.pictures(source)], ["SACCR_Results_EAD.png"])
            self.assertEqual(screenshots.main([str(source), "--root", str(root), "--write"]), 0)
            written = root / "docs" / "assets" / "SACCR_Results_EAD.png"
            with Image.open(written) as image:
                self.assertEqual(image.mode, "P")
                self.assertEqual(image.size, (40, 20))
                self.assertLessEqual(len(image.getcolors(256) or []), 256)
            self.assertFalse((root / "docs" / "assets" / "busy-dialog-1.png").exists())

    def test_a_folder_without_pictures_fails(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            self.assertEqual(screenshots.main([folder]), 1)


if __name__ == "__main__":
    unittest.main()
