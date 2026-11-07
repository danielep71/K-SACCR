"""The sheet navigation bar of the workbook template (#615): tools/navigation_bar.py."""
import re
import shutil
import tempfile
import unittest
from pathlib import Path

import navigation_bar as bar

ROOT = Path(__file__).resolve().parents[1]


class NavigationBarTests(unittest.TestCase):
    def setUp(self) -> None:
        self.folder = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, self.folder)
        for path in (bar.TEMPLATE, bar.CONTRACT):
            (self.folder / path).parent.mkdir(parents=True, exist_ok=True)
            shutil.copyfile(ROOT / path, self.folder / path)

    def contract(self) -> Path:
        return self.folder / bar.CONTRACT

    def test_committed_template_passes(self) -> None:
        self.assertEqual(bar.problems(ROOT), [])

    def test_write_is_idempotent(self) -> None:
        bar.write(self.folder)
        first = (self.folder / bar.TEMPLATE).read_bytes()
        bar.write(self.folder)
        self.assertEqual((self.folder / bar.TEMPLATE).read_bytes(), first)
        self.assertEqual(bar.problems(self.folder), [])

    def test_blocks_put_inputs_on_top_and_home_last(self) -> None:
        top, rest = bar.blocks(bar.sections(self.contract().read_text(encoding="utf-8")))
        self.assertEqual([name for name, _ in top[0][1]], ["Params", "NettingSets", "Trades"])
        self.assertEqual(rest[-1][1][-1][0], "Home")

    def test_top_block_fits_the_frozen_rows_of_every_sheet(self) -> None:
        contract = self.contract().read_text(encoding="utf-8")
        top, _ = bar.blocks(bar.sections(contract))
        package = bar.Package(ROOT / bar.TEMPLATE)
        for name, part in bar.sheets_of(package):
            room = bar.row_offset_points(package.text(part), bar.TOP_ROWS) / 0.75
            self.assertGreaterEqual(room, bar.block_height(top), name)

    def test_renamed_tooltip_makes_the_template_stale(self) -> None:
        text = self.contract().read_text(encoding="utf-8")
        self.contract().write_text(text.replace('supertip="One row per trade leg."',
                                                'supertip="One row per leg."'), encoding="utf-8")
        self.assertTrue(any("not current" in p for p in bar.problems(self.folder)))

    def test_menu_that_misses_a_sheet_fails(self) -> None:
        text = self.contract().read_text(encoding="utf-8")
        text = re.sub(r'\s*<button id="ksGoBuckets"[^>]*/>', "", text)
        self.contract().write_text(text, encoding="utf-8")
        self.assertTrue(any("Go to menu lists" in p for p in bar.problems(self.folder)))

    def test_template_without_the_bar_fails(self) -> None:
        shutil.copyfile(self.folder / bar.TEMPLATE, self.folder / "copy.xlsx")
        package = bar.Package(self.folder / bar.TEMPLATE)
        for name in list(package.entries):
            if name.startswith("xl/drawings/drawing") and name.endswith(".xml"):
                drawing = package.text(name)
                rels = bar.rels_path(name)
                stripped, kept = bar.strip_bar(drawing, package.text(rels) if package.has(rels) else bar.EMPTY_RELS)
                package.put(name, stripped)
                if package.has(rels):
                    package.put(rels, kept)
        package.write(self.folder / bar.TEMPLATE)
        found = bar.problems(self.folder)
        self.assertTrue(any("NavBarTop missing" in p for p in found))
        self.assertTrue(any("not current" in p for p in found))

    def test_missing_contract_fails(self) -> None:
        self.contract().unlink()
        self.assertEqual(bar.problems(self.folder), [f"{bar.CONTRACT} is missing"])


if __name__ == "__main__":
    unittest.main()
