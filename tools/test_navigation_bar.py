"""The sheet navigation bar of the workbook template (#615): tools/navigation_bar.py."""
import posixpath
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
        for path in (bar.TEMPLATE, bar.CONTRACT, bar.CONFIG):
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

    def test_blocks_put_home_and_the_inputs_on_top(self) -> None:
        # The owner's request of 2026-11-26: Home at the top of the bar, above Inputs.
        top, rest = bar.blocks(bar.sections(self.contract().read_text(encoding="utf-8")))
        self.assertEqual([[name for name, _ in items] for _, items in top],
                         [["Home"], ["Params", "NettingSets", "Trades"]])
        self.assertEqual([title for title, _ in rest], ["Staging", "Outputs", "Reconciliation", "Reference and tests"])

    def test_top_block_fits_the_frozen_rows_of_every_sheet(self) -> None:
        contract = self.contract().read_text(encoding="utf-8")
        top, _ = bar.blocks(bar.sections(contract))
        package = bar.Package(ROOT / bar.TEMPLATE)
        for name, part in bar.sheets_of(package):
            xml = package.text(part)
            frozen = int(re.search(r'<pane [^>]*ySplit="(\d+)"', xml).group(1))  # type: ignore[union-attr]
            room = bar.row_offset_points(xml, frozen) / 0.75
            self.assertGreaterEqual(room, bar.block_height(top), name)

    def test_drawing_the_bar_raises_no_row(self) -> None:
        # Home fits in the frozen rows as they are: no row of any sheet changes height.
        def heights(package: "bar.Package") -> dict[str, list[str]]:
            return {name: re.findall(r'<row r="\d+"[^>]*?\bht="[\d.]+"', package.text(part))
                    for name, part in bar.sheets_of(package)}

        before = heights(bar.Package(self.folder / bar.TEMPLATE))
        bar.write(self.folder)
        self.assertEqual(heights(bar.Package(self.folder / bar.TEMPLATE)), before)

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

    def test_config_group_name_must_match(self) -> None:
        config = self.folder / bar.CONFIG
        text = config.read_text(encoding="cp1252")
        config.write_text(text.replace('NAV_GROUP_REST  As String = "NavBar"', 'NAV_GROUP_REST  As String = "Nav"'),
                          encoding="cp1252")
        self.assertTrue(any("NAV_GROUP_REST" in p for p in bar.problems(self.folder)))

    def drawings(self, package: "bar.Package") -> dict[str, str]:
        """The drawing of every sheet, by sheet name."""
        found = {}
        for name, part in bar.sheets_of(package):
            sheet_rels = package.text(bar.rels_path(part))
            rid = re.search(r'<drawing r:id="([^"]+)"/>', package.text(part)).group(1)  # type: ignore[union-attr]
            target = bar.relationships(sheet_rels)[rid][1]
            found[name] = package.text(posixpath.normpath(posixpath.join(posixpath.dirname(part), target)))
        return found

    def test_every_sheet_has_one_triangle_shown_at_the_far_left_of_b1(self) -> None:
        # #777: one triangle at a time, in the same place, both running the toggle: the one that hides the
        # bar shown, the one that shows it saved hidden.
        package = bar.Package(ROOT / bar.TEMPLATE)
        drawings = self.drawings(package)
        self.assertEqual(len(drawings), 21)
        for name, drawing in drawings.items():
            anchors = {n: a for a in bar.ANCHOR.findall(drawing)
                       for n in (bar.TOGGLE_HIDE, bar.TOGGLE_SHOW) if 'name="%s"' % n in a}
            with self.subTest(sheet=name):
                self.assertEqual(sorted(anchors), sorted((bar.TOGGLE_HIDE, bar.TOGGLE_SHOW)))
                for toggle, anchor in anchors.items():
                    self.assertIn('macro="[0]!ToggleNavigationBar"', anchor)
                    self.assertIn("<xdr:col>1</xdr:col><xdr:colOff>%d</xdr:colOff><xdr:row>0</xdr:row>"
                                  % (bar.TOGGLE_X * bar.PX), anchor)
                    self.assertIn('fPrintsWithSheet="0"', anchor)
                    self.assertEqual('hidden="1"' in anchor, toggle == bar.TOGGLE_SHOW)
                # The one that hides the bar points left: its tip is on the left edge.
                self.assertIn('<a:lnTo><a:pt x="0" y="%d"/>' % (bar.TOGGLE_H // 2), anchors[bar.TOGGLE_HIDE])
                self.assertIn('<a:lnTo><a:pt x="%d" y="%d"/>' % (bar.TOGGLE_W, bar.TOGGLE_H // 2),
                              anchors[bar.TOGGLE_SHOW])

    def test_the_triangle_fits_in_column_b_and_the_titles_sit_in_column_c(self) -> None:
        self.assertLessEqual(bar.TOGGLE_X + bar.TOGGLE_W, bar.GUTTER_PX - 2)
        package = bar.Package(ROOT / bar.TEMPLATE)
        styles = package.text("xl/styles.xml")
        xfs = re.findall(r"<xf\b[^>]*?(?:/>|>.*?</xf>)",
                         re.search(r"<cellXfs[^>]*>(.*?)</cellXfs>", styles, re.S).group(1), re.S)  # type: ignore[union-attr]
        for name, part in bar.sheets_of(package):
            xml = package.text(part)
            self.assertIn('<col min="1" max="1" width="18" ', xml)
            self.assertIn('<col min="2" max="2" width="2" customWidth="1"/>', xml)
            for row in bar.TITLE_ROWS:
                margin = bar.cell_element(xml, "B%d" % row)
                title = bar.cell_element(xml, "C%d" % row)
                with self.subTest(sheet=name, row=row):
                    assert margin is not None and title is not None
                    self.assertNotIn("<v>", margin.group(0))
                    self.assertIn("<v>", title.group(0))
                    style = int(re.search(r' s="(\d+)"', title.group(0)).group(1))  # type: ignore[union-attr]
                    self.assertNotIn("indent=", xfs[style])
                    self.assertIn('horizontal="left"', xfs[style])

    def test_a_title_in_b1_is_moved_to_c1(self) -> None:
        # A template from before #777: the title in B1, the band's empty cell in C1.
        package = bar.Package(self.folder / bar.TEMPLATE)
        _, part = bar.sheets_of(package)[0]
        xml = package.text(part)
        title, margin = bar.cell_element(xml, "C1"), bar.cell_element(xml, "B1")
        assert title is not None and margin is not None
        swapped = (xml[:margin.start()] + title.group(0).replace('r="C1"', 'r="B1"')
                   + xml[margin.end():title.start()] + margin.group(0).replace('r="B1"', 'r="C1"') + xml[title.end():])
        package.put(part, swapped)
        package.write(self.folder / bar.TEMPLATE)
        self.assertTrue(any("not in C1" in p for p in bar.problems(self.folder)))
        bar.write(self.folder)
        self.assertEqual((self.folder / bar.TEMPLATE).read_bytes(), (ROOT / bar.TEMPLATE).read_bytes())

    def test_columns_a_and_b_are_frozen_on_every_sheet(self) -> None:
        package = bar.Package(ROOT / bar.TEMPLATE)
        for name, part in bar.sheets_of(package):
            pane = re.search(r'<pane xSplit="(\d+)"', package.text(part))
            with self.subTest(sheet=name):
                self.assertIsNotNone(pane)
                self.assertGreaterEqual(int(pane.group(1)), 2)  # type: ignore[union-attr]

    def test_config_triangle_name_must_match(self) -> None:
        config = self.folder / bar.CONFIG
        text = config.read_text(encoding="cp1252")
        config.write_text(text.replace('"NavBarShow"', '"NavShow"'), encoding="cp1252")
        self.assertTrue(any("NAV_TOGGLE_SHOW" in p for p in bar.problems(self.folder)))

    def test_missing_contract_fails(self) -> None:
        self.contract().unlink()
        self.assertEqual(bar.problems(self.folder), [f"{bar.CONTRACT} is missing"])


if __name__ == "__main__":
    unittest.main()
