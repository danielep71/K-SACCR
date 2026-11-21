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
                         ["SF_Key", "Currency Code", "Bank Holiday", "Description"])

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

    def test_the_fx_table_has_no_pair_or_price_column_and_the_engine_does_not_read_prices(self) -> None:
        # Download prices builds the pair in code and keeps its price formulas in a hidden scratch column.
        for row in (60, 61, 62, 81, 82, 260):
            for column in "FG":
                self.assertIsNone(self.sheet.find(f'.//{MAIN}c[@r="{column}{row}"]'), (column, row))
        self.assertNotIn("$C$61&amp;", self.xml)
        engine = (ROOT / "src" / "core" / "CORE_Engine.bas").read_text(encoding="cp1252")
        self.assertNotIn("Refinitiv", engine)
        self.assertNotIn("STOCKHISTORY", engine)

    def test_download_prices_writes_the_rates_and_leaves_no_formula_in_the_sheet(self) -> None:
        # A STOCKHISTORY formula left in the sheet is a data request that Excel may never finish: a run
        # in manual calculation then refuses to start. The macro writes the formulas into a hidden scratch
        # column, reads the prices, writes them as the rates and empties the column.
        self.assertNotIn("STOCKHISTORY", self.xml)
        for row in range(61, 261):
            self.assertIsNone(self.sheet.find(f'.//{MAIN}c[@r="AR{row}"]'), row)
        main = (ROOT / "src" / "workbook" / "M_Main.bas").read_text(encoding="cp1252")
        self.assertIn('"LET(r,STOCKHISTORY("""', main)
        for form in ('Base & ":" & Quote', 'Base & "/" & Quote', "Base & Quote"):  # the three names of a pair
            self.assertIn(form, main)
        self.assertIn("Application.Cursor = xlWait", main)  # the wait shows
        self.assertNotIn("from STOCKHISTORY?", main)  # the confirmation does not name the function
        self.assertIn('CallByName PriceCell, "Formula2", VbLet', main)  # late bound: it compiles in any Excel
        self.assertRegex(main, r"(?m)^Public Sub DownloadClosingPrices\(\)")
        self.assertRegex(main, r"(?m)^Public Function DownloadClosingPrices_Silent\(\) As String")
        for removed in ("CopyClosingPrices", "ScanClosingPrices", "FX_COL_PRICE_PAIR"):
            self.assertNotIn(removed, main)

    def test_the_fx_table_is_filled_with_twenty_currencies_against_the_euro(self) -> None:
        codes = [self.text(f"C{row}") for row in range(61, 82)]
        self.assertEqual(codes[0], "EUR")
        self.assertEqual(len(set(codes)), 21)
        self.assertTrue(all(re.fullmatch(r"[A-Z]{3}", code) for code in codes), codes)
        for row in range(61, 82):
            self.assertTrue(self.text(f"D{row}"), row)
            self.assertGreater(float(self.text(f"E{row}")), 0, row)
        self.assertEqual(self.text("E61"), "1")  # the reporting currency against itself
        self.assertEqual(len({self.text(f"E{row}") for row in range(62, 82)}), 20)  # no rate is a copied placeholder
        self.assertTrue(self.text("D82").startswith("Add currencies directly below the last row"))
        self.assertEqual(self.text("C82"), "")  # a note in column C would be read as a currency

    def test_the_parameters_are_in_the_owners_order_with_their_names(self) -> None:
        order = [self.text(f"C{row}") for row in (8, 9, 10, 23, 24, 25, 26, 27)]
        self.assertEqual(order, ["Regime", "Alpha", "MultiplierFloor", "LambdaThresholdIR", "LambdaThresholdCO",
                                 "DaycountConvention", "YearConvention", "Calendar"])
        for name, cell in (("Regime", "$E$8"), ("Alpha", "$E$9"), ("MultiplierFloor", "$E$10"), ("LambdaThresholdIR", "$E$23"),
                           ("LambdaThresholdCO", "$E$24"), ("DaycountConvention", "$E$25"), ("YearConvention", "$E$26"),
                           ("Calendar", "$E$27")):
            self.assertIn(f'<definedName name="{name}">Params!{cell}</definedName>', self.workbook, name)
        for old in ("DayCountConvention", "DayCountNumerator"):
            self.assertNotIn(f'name="{old}"', self.workbook)
        # a validation belongs to the cell it checks: the regime list on E8, the currency check on E7 only
        self.assertRegex(self.xml, r'<dataValidation [^>]*sqref="E8"[^>]*><formula1>"CRR,BCBS"</formula1>')
        self.assertRegex(self.xml, r'<dataValidation [^>]*sqref="E7"')
        self.assertNotRegex(self.xml, r'sqref="E7:E8"')
        self.assertRegex(self.xml, r'<dataValidation [^>]*sqref="E25"[^>]*><formula1>"ACTUAL"</formula1>')

    def test_the_rate_header_is_ten_points_as_the_style_rules_allow(self) -> None:
        styles = ElementTree.fromstring(self.package.read("xl/styles.xml"))
        xfs = list(styles.find(MAIN + "cellXfs"))  # type: ignore[arg-type]
        fonts = list(styles.find(MAIN + "fonts"))  # type: ignore[arg-type]
        cell = self.sheet.find(f'.//{MAIN}c[@r="E60"]')
        assert cell is not None
        font = fonts[int(xfs[int(cell.attrib["s"])].attrib["fontId"])]
        size = font.find(MAIN + "sz")
        assert size is not None
        self.assertEqual(size.attrib["val"], "10")
        rules = (ROOT / "docs" / "workbook" / "sheet-style.json").read_text(encoding="utf-8")
        self.assertIn('"header_size_exceptions"', rules)
        self.assertIn('"Params": {"E60": 10}', rules)

    def test_the_date_of_the_prices_is_shown_beside_the_button(self) -> None:
        self.assertIn('<definedName name="FX_PriceDate">Params!$AQ$27</definedName>', self.workbook)
        self.assertEqual(self.text("AQ27"), self.text("E6"))  # the date of the prices in the template: the reporting date
        self.assertEqual(self.text("G59"), "Prices downloaded for 2026-09-30")
        formula = self.sheet.find(f'.//{MAIN}c[@r="G59"]/{MAIN}f')
        assert formula is not None and formula.text is not None
        self.assertIn("FX_PriceDate", formula.text)
        self.assertIn('TEXT(MONTH(FX_PriceDate),"00")', formula.text)  # no locale-bound date format in the formula
        hidden = [(int(col.attrib["min"]), int(col.attrib["max"])) for col in self.sheet.iter(MAIN + "col")
                  if col.attrib.get("hidden") == "1"]
        self.assertTrue(all(any(low <= column <= high for low, high in hidden) for column in (43, 44)), hidden)  # AQ, AR
        config = (ROOT / "src" / "core" / "CORE_Config.bas").read_text(encoding="cp1252")
        self.assertIn('NM_FX_PRICE_DATE As String = "FX_PriceDate"', config)
        self.assertIn("FX_COL_PRICE As Long = 41", config)  # the scratch column AR is 41 columns right of the key column C
        self.assertNotIn("FX_COL_PRICE_PAIR", config)

    def test_a_currency_appended_below_the_table_has_a_pair_from_its_code(self) -> None:
        main = (ROOT / "src" / "workbook" / "M_Main.bas").read_text(encoding="cp1252")
        # the pair is built from the code and the reporting currency in code, for every row down to the first blank key
        self.assertIn("ClosingPair = Reporting & \"/\" & Code", main)

    def test_the_closing_day_validation_and_the_unlocked_cells_follow_the_table(self) -> None:
        self.assertRegex(self.xml, r'sqref="I61:I160"[^>]*><formula1>AND\(ISNUMBER\(I61\),I61&gt;=61\)</formula1>')
        self.assertNotRegex(self.xml, r'sqref="G\d+:G\d+"[^>]*><formula1>AND\(ISNUMBER\(G')
        matrix = (ROOT / "docs" / "workbook" / "protection.json").read_text(encoding="utf-8")
        self.assertIn('"C61:E260", "I61:J160"', matrix)
        self.assertIn('"unlocked_cells": ["E6", "E7", "E8", "E26", "E27"', matrix)

    def test_the_trades_validations_read_the_moved_tables(self) -> None:
        trades = self.package.read(sheet_parts(self.package)["Trades"]).decode("utf-8")
        self.assertIn("Params!$E$32:$E$57", trades)
        self.assertEqual(sorted(set(re.findall(r"Params!\$C\$61:\$C\$81", trades))), ["Params!$C$61:$C$81"])
        self.assertNotIn("Params!$C$60:$C$79", trades)

    def test_the_fx_table_holds_the_euro_and_twenty_currencies(self) -> None:
        # rows 61 to 81: the reporting currency and 20 others, validated and checked for a repeated code
        self.assertRegex(self.xml, r'sqref="C61:C81"[^>]*><formula1>AND\(ISTEXT\(C61\),LEN\(C61\)=3,COUNTIF\(\$C\$61:\$C\$81,C61\)=1\)')
        self.assertRegex(self.xml, r'<dataValidation [^>]*sqref="E61:E81"')
        self.assertRegex(self.xml, r'<conditionalFormatting sqref="C61:C81">')
        self.assertNotIn('ReportingCcy&amp;"/"&amp;C', self.xml)  # no pair formula: the macro builds the pair


if __name__ == "__main__":
    unittest.main()
