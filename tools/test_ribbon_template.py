"""The Ribbon gate of the workbook template (#214): tools/ribbon_template.py."""
import shutil
import tempfile
import unittest
import zipfile
from pathlib import Path

import ribbon_template as gate

ROOT = Path(__file__).resolve().parents[1]


class RibbonTemplateTests(unittest.TestCase):
    def setUp(self) -> None:
        self.folder = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, self.folder)
        for path in (gate.TEMPLATE, gate.SOURCE):
            (self.folder / path).parent.mkdir(parents=True, exist_ok=True)
            shutil.copyfile(ROOT / path, self.folder / path)
        shutil.copytree(ROOT / gate.ICONS, self.folder / gate.ICONS)

    def rewrite(self, change: dict[str, bytes | None]) -> None:
        template = self.folder / gate.TEMPLATE
        with zipfile.ZipFile(template) as package:
            entries = [(info, package.read(info.filename)) for info in package.infolist()]
        with zipfile.ZipFile(template, "w") as output:
            for info, data in entries:
                if info.filename in change:
                    replacement = change[info.filename]
                    if replacement is None:
                        continue
                    data = replacement
                output.writestr(info, data, compress_type=info.compress_type)

    def test_committed_template_passes(self) -> None:
        self.assertEqual(gate.problems(ROOT), [])

    def test_stale_part_fails(self) -> None:
        source = self.folder / gate.SOURCE
        source.write_bytes(source.read_bytes().replace(b'label="K-SACCR"', b'label="K-SACCR changed"'))
        self.assertTrue(any("differs" in p for p in gate.problems(self.folder)))

    def test_missing_part_fails(self) -> None:
        self.rewrite({gate.PART: None})
        self.assertTrue(any("no customUI" in p for p in gate.problems(self.folder)))

    def test_missing_relationship_fails(self) -> None:
        with zipfile.ZipFile(self.folder / gate.TEMPLATE) as package:
            rels = package.read(gate.RELATIONSHIPS).decode("utf-8")
        self.rewrite({gate.RELATIONSHIPS: rels.replace(gate.RELATIONSHIP, "").encode("utf-8")})
        self.assertTrue(any("relationship" in p for p in gate.problems(self.folder)))

    def test_missing_source_fails(self) -> None:
        (self.folder / gate.SOURCE).unlink()
        self.assertEqual(gate.problems(self.folder), [f"{gate.SOURCE} is missing"])

    def test_write_repairs_and_is_idempotent(self) -> None:
        source = self.folder / gate.SOURCE
        source.write_bytes(source.read_bytes().replace(b'label="K-SACCR"', b'label="K-SACCR changed"'))
        gate.write(self.folder)
        self.assertEqual(gate.problems(self.folder), [])
        first = (self.folder / gate.TEMPLATE).read_bytes()
        gate.write(self.folder)
        self.assertEqual((self.folder / gate.TEMPLATE).read_bytes(), first)
        with zipfile.ZipFile(self.folder / gate.TEMPLATE) as package:
            rels = package.read(gate.RELATIONSHIPS).decode("utf-8")
        self.assertEqual(gate.ribbon_relationships(rels), [gate.RELATIONSHIP])

    def test_the_template_carries_the_icons_the_source_names(self) -> None:
        # #803: each image="NAME" of the source is customUI/images/NAME.png with an image relationship.
        names = gate.image_names((ROOT / gate.SOURCE).read_bytes())
        self.assertGreaterEqual(len(names), 12)
        with zipfile.ZipFile(ROOT / gate.TEMPLATE) as package:
            rels = package.read(gate.PART_RELATIONSHIPS).decode("utf-8")
            for name in names:
                with self.subTest(image=name):
                    self.assertEqual(package.read(gate.IMAGE_FOLDER + name + ".png"),
                                     (ROOT / gate.ICONS / f"{name}.png").read_bytes())
                    self.assertIn(f'<Relationship Id="{name}" Type="{gate.IMAGE_TYPE}" '
                                  f'Target="images/{name}.png"/>', rels)

    def test_a_missing_or_changed_icon_fails(self) -> None:
        self.assertEqual(gate.problems(self.folder), [])
        self.rewrite({gate.IMAGE_FOLDER + "ksiRunLarge.png": b"other"})
        self.assertTrue(any("ksiRunLarge.png differs" in p for p in gate.problems(self.folder)))
        self.rewrite({gate.IMAGE_FOLDER + "ksiRunLarge.png": None})
        self.assertTrue(any("no customUI/images/ksiRunLarge.png" in p for p in gate.problems(self.folder)))
        (self.folder / gate.ICONS / "ksiHomeLarge.png").unlink()
        self.assertTrue(any("ksiHomeLarge.png is missing" in p for p in gate.problems(self.folder)))

    def test_a_left_over_customui_part_fails(self) -> None:
        with zipfile.ZipFile(self.folder / gate.TEMPLATE, "a") as package:
            package.writestr("customUI/images/old.png", b"x")
        self.assertTrue(any("unexpected customUI part customUI/images/old.png" in p
                            for p in gate.problems(self.folder)))


if __name__ == "__main__":
    unittest.main()
