"""Tests of the layout of the Params sheet of the workbook template (#721).

The day-count convention is chosen from a list kept in hidden cells; the FX table carries the currency
pair and a closing-price column that the engine does not read; the table of closing days sits right of
them, and the names, the validations and the editable cells follow the layout.
"""

from __future__ import annotations

import re
import unittest
import zipfile
from pathlib import Path
from xml.etree import ElementTree

from test_template_buttons import sheet_parts

ROOT = Path(__file__).resolve().parents[1]
TEMPLATE = ROOT / "src" / "workbook" / "SACCR_Template.xlsx"
MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
BASES = ["366", "365", "364", "360", "252", "250"]


class ParamsLayoutTests(unittest.TestCase):
    def setUp(self) -> None:
        self.package = zipfile.ZipFile(TEMPLATE)
        self.addCleanup(self.package.close)
        self.xml = self.package.read(sheet_parts(self.package)["Params"]).decode("utf-8")
        self.sheet = ElementTree.fromstring(self.xml)
        self.workbook = self.package.read("xl/workbook.xml").decode("utf-8")
        self.strings = [re.sub(r"<[^>]+>", "", item) for item in
                        re.findall(r"<si>(.*?)</si>", self.package.read("xl/sharedStrings.xml").decode("utf-8"), re.S)]

    def text(self, ref: str) -> str:
        cell = self.sheet.find(f'.//{MAIN}c[@r="{ref}"]')
        if cell is None:
            return ""
        value = cell.findtext(MAIN + "v") or ""
        return self.strings[int(value)] if cell.attrib.get("t") == "s" and value else value

    def test_the_tables_start_where_the_names_say(self) -> None:
        for name, target in (("SF_Table", "$C$31:$K$57"), ("FX_Table", "$C$60:$E$65"), ("Holiday_Table", "$I$60:$J$160")):
            self.assertIn(f'<definedName name="{name}">Params!{target}</definedName>', self.workbook)
        self.assertEqual([self.text(ref) for ref in ("C31", "C60", "I60", "J60")],
                         ["SF_Key", "FX_Ccy", "Closing day", "Description"])

    def test_the_day_count_list_is_in_hidden_cells_and_feeds_the_choice(self) -> None:
        self.assertEqual([self.text(f"AQ{row}") for row in range(20, 26)], BASES)
        self.assertEqual(self.text("AQ26"), "")
        self.assertEqual([self.text(f"M{row}") for row in range(20, 27)], [""] * 7)
        self.assertRegex(self.xml, r'<dataValidation [^>]*sqref="E26"[^>]*><formula1>\$AQ\$20:\$AQ\$25</formula1>')
        self.assertIn("Choose 366, 365, 364, 360, 252 or 250.", self.xml)
        hidden = [(int(col.attrib["min"]), int(col.attrib["max"])) for col in self.sheet.iter(MAIN + "col")
                  if col.attrib.get("hidden") == "1"]
        self.assertTrue(any(low <= 43 <= high for low, high in hidden), hidden)  # AQ is column 43

    def test_the_engine_accepts_the_same_denominators(self) -> None:
        engine = (ROOT / "src" / "core" / "CORE_Engine.bas").read_text(encoding="cp1252")
        self.assertIn("Case 366#, 365#, 364#, 360#, 252#, 250#", engine)
        config = (ROOT / "src" / "core" / "CORE_Config.bas").read_text(encoding="cp1252")
        self.assertIn('"366, 365, 364, 360, 252 or 250"', config)

    def test_the_calendar_note_runs_over_two_rows(self) -> None:
        self.assertTrue(self.text("F27").endswith("so an empty table gives calendar days."))
        self.assertTrue(self.text("F28").startswith("TARGET: the business days"))
        self.assertIn("closed too. Business days are normally used only with a denominator of 252 or 250.", self.text("F28"))
        self.assertEqual(self.text("C30"), "Supervisory parameters (CRE52.72 / CRR Art. 279a, 280a-280f)")

    def test_the_fx_table_has_a_pair_and_a_closing_price_the_engine_does_not_read(self) -> None:
        self.assertEqual((self.text("F60"), self.text("G60")), ("Currency Pair", "Closing Price\nfrom Refinitiv"))
        for row, pair in ((62, "USD"), (63, "GBP"), (64, "JPY"), (65, "CHF")):
            self.assertIn(f'$C$61&amp;"/"&amp;C{row}', self.xml)
            self.assertEqual(self.text(f"F{row}"), f"EUR/{pair}")
            self.assertIn(f"_xlfn.STOCKHISTORY(F{row},AsOfDate,AsOfDate,0,0,1)", self.xml)
        engine = (ROOT / "src" / "core" / "CORE_Engine.bas").read_text(encoding="cp1252")
        self.assertNotIn("Refinitiv", engine)
        self.assertNotIn("STOCKHISTORY", engine)

    def test_the_closing_day_validation_and_the_unlocked_cells_follow_the_table(self) -> None:
        self.assertRegex(self.xml, r'sqref="I61:I160"[^>]*><formula1>AND\(ISNUMBER\(I61\),I61&gt;=61\)</formula1>')
        self.assertNotRegex(self.xml, r'sqref="G\d+:G\d+"[^>]*><formula1>AND\(ISNUMBER\(G')
        matrix = (ROOT / "docs" / "workbook" / "protection.json").read_text(encoding="utf-8")
        self.assertIn('"C61:E260", "I61:J160"', matrix)

    def test_the_trades_validations_read_the_moved_tables(self) -> None:
        trades = self.package.read(sheet_parts(self.package)["Trades"]).decode("utf-8")
        self.assertIn("Params!$E$32:$E$57", trades)
        self.assertEqual(sorted(set(re.findall(r"Params!\$C\$61:\$C\$80", trades))), ["Params!$C$61:$C$80"])
        self.assertNotIn("Params!$C$60:$C$79", trades)


if __name__ == "__main__":
    unittest.main()
