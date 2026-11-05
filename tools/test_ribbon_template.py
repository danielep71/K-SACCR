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
        source.write_bytes(source.read_bytes().replace(b'label="SA-CCR"', b'label="SA-CCR changed"'))
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
        source.write_bytes(source.read_bytes().replace(b'label="SA-CCR"', b'label="SA-CCR changed"'))
        gate.write(self.folder)
        self.assertEqual(gate.problems(self.folder), [])
        first = (self.folder / gate.TEMPLATE).read_bytes()
        gate.write(self.folder)
        self.assertEqual((self.folder / gate.TEMPLATE).read_bytes(), first)
        with zipfile.ZipFile(self.folder / gate.TEMPLATE) as package:
            rels = package.read(gate.RELATIONSHIPS).decode("utf-8")
        self.assertEqual(gate.ribbon_relationships(rels), [gate.RELATIONSHIP])


if __name__ == "__main__":
    unittest.main()
