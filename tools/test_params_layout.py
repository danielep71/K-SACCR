"""Tests of the layout of the Params sheet of the workbook template (#721).

The day-count convention is chosen from a list kept in hidden cells; the FX table carries the currency
pair and a closing-price column that the engine does not read; the table of closing days sits right of
them, and the names, the validations and the editable cells follow the layout. The source of the rates
Download prices writes is chosen beside the button and stated with their date (#745). Every position is
read from a workbook name (#756): the owner may move rows in Excel, which keeps the names on their cells.
"""

from __future__ import annotations

import re
import unittest
import zipfile
from pathlib import Path
from xml.etree import ElementTree

from template_layout import NamedCell, named_cell
from test_template_buttons import sheet_parts

ROOT = Path(__file__).resolve().parents[1]
TEMPLATE = ROOT / "src" / "workbook" / "SACCR_Template.xlsx"
MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
BASES = ["366", "365", "364", "360", "252", "250"]
FX_ROWS = 21  # the reporting currency and twenty others
FX_SPARE = 200  # rows of the FX table, its currencies and the rows below them for added currencies
HOLIDAY_ROWS = 100


class ParamsLayoutTests(unittest.TestCase):
    def setUp(self) -> None:
        self.package = zipfile.ZipFile(TEMPLATE)
        self.addCleanup(self.package.close)
        self.xml = self.package.read(sheet_parts(self.package)["Params"]).decode("utf-8")
        self.sheet = ElementTree.fromstring(self.xml)
        self.workbook = self.package.read("xl/workbook.xml").decode("utf-8")
        self.strings = [re.sub(r"<[^>]+>", "", item) for item in
                        re.findall(r"<si>(.*?)</si>", self.package.read("xl/sharedStrings.xml").decode("utf-8"), re.S)]

    def at(self, name: str) -> NamedCell:
        cell = named_cell(self.package, name)
        self.assertEqual(cell.sheet, "Params", name)
        return cell

    def row(self, name: str) -> int:
        return self.at(name).row

    def validation(self, ref: str) -> str:
        """The first formula of the data validation whose range starts at ref."""
        match = re.search(rf'<dataValidation [^>]*sqref="{ref}(?::[A-Z]+\d+)?"[^>]*>.*?<formula1>(.*?)</formula1>',
                          self.xml, re.S)
        self.assertIsNotNone(match, ref)
        return match.group(1) if match else ""

    def text(self, ref: str) -> str:
        cell = self.sheet.find(f'.//{MAIN}c[@r="{ref}"]')
        if cell is None:
            return ""
        value = cell.findtext(MAIN + "v") or ""
        return self.strings[int(value)] if cell.attrib.get("t") == "s" and value else value

    def test_the_tables_start_where_the_names_say(self) -> None:
        sf, fx, holiday = self.at("SF_Table"), self.at("FX_Table"), self.at("Holiday_Table")
        self.assertEqual([self.text(f"{sf.column}{sf.row}"), self.text(f"{fx.column}{fx.row}"),
                          self.text(f"{holiday.column}{holiday.row}"), self.text(f"J{holiday.row}")],
                         ["SF_Key", "Currency Code", "Bank Holiday", "Description"])
        self.assertEqual((sf.column, fx.column, holiday.column), ("C", "C", "I"))
        self.assertLess(sf.row, fx.row)
        self.assertEqual(fx.row, holiday.row)  # the closing days sit right of the FX table, on its header row
        for name, last_column, rows in (("SF_Table", "K", 26), ("FX_Table", "E", 5), ("Holiday_Table", "J", HOLIDAY_ROWS)):
            first = self.row(name)
            self.assertIn(f'<definedName name="{name}">Params!${self.at(name).column}${first}:${last_column}${first + rows}'
                          "</definedName>", self.workbook, name)

    def test_the_day_count_list_is_in_hidden_cells_and_feeds_the_choice(self) -> None:
        source = re.fullmatch(r"\$AQ\$(\d+):\$AQ\$(\d+)", self.validation(f"E{self.row('YearConvention')}"))
        assert source is not None
        first, last = int(source.group(1)), int(source.group(2))
        self.assertEqual([self.text(f"AQ{row}") for row in range(first, last + 1)], BASES)
        self.assertEqual(self.text(f"AQ{last + 1}"), "")
        self.assertEqual([self.text(f"M{row}") for row in range(first, last + 2)], [""] * (len(BASES) + 1))
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
        calendar = self.row("Calendar")
        self.assertTrue(self.text(f"F{calendar}").endswith("so an empty table gives calendar days."))
        self.assertTrue(self.text(f"F{calendar + 1}").startswith("TARGET: the business days"))
        self.assertIn("closed too. Business days are normally used only with a denominator of 252 or 250.",
                      self.text(f"F{calendar + 1}"))
        self.assertEqual(self.text(f"C{self.row('SF_Table') - 1}"),
                         "Supervisory parameters (CRE52.72 / CRR Art. 279a, 280a-280f)")

    def test_the_fx_table_has_no_pair_or_price_column_and_the_engine_does_not_read_prices(self) -> None:
        # Download prices builds the pair in code and keeps its price formulas in a hidden scratch column.
        header = self.row("FX_Table")
        for row in (header, header + 1, header + 2, header + FX_ROWS, header + FX_ROWS + 1, header + FX_SPARE):
            for column in "FG":
                self.assertIsNone(self.sheet.find(f'.//{MAIN}c[@r="{column}{row}"]'), (column, row))
        self.assertNotIn(f"$C${header + 1}&amp;", self.xml)
        engine = (ROOT / "src" / "core" / "CORE_Engine.bas").read_text(encoding="cp1252")
        self.assertNotIn("Refinitiv", engine)
        self.assertNotIn("STOCKHISTORY", engine)

    def test_download_prices_writes_the_rates_and_leaves_no_formula_in_the_sheet(self) -> None:
        # A STOCKHISTORY formula left in the sheet is a data request that Excel may never finish: a run
        # in manual calculation then refuses to start. The macro writes the formulas into a hidden scratch
        # column, reads the prices, writes them as the rates and empties the column.
        self.assertNotIn("STOCKHISTORY", self.xml)
        header = self.row("FX_Table")
        for row in range(header + 1, header + FX_SPARE + 1):
            self.assertIsNone(self.sheet.find(f'.//{MAIN}c[@r="AR{row}"]'), row)
        main = (ROOT / "src" / "workbook" / "M_Main.bas").read_text(encoding="cp1252")
        self.assertIn('"LET(r,STOCKHISTORY("""', main)
        self.assertIn('Replace(Pair, "/", ":")', main)  # STOCKHISTORY takes EUR:USD
        self.assertIn("PricesPending(", main)  # waits while a cell shows #BUSY!
        # Excel hands the data back only when no macro runs: the button returns and a timer reads the prices
        self.assertIn("Application.OnTime mPollAt, mPollName", main)
        self.assertRegex(main, r"(?m)^Public Sub PollClosingPrices\(\)")
        self.assertNotIn("CalculateUntilAsyncQueriesDone", main)
        self.assertNotIn("CVErr(2045)", main)  # 2045 is #SPILL!, not the busy state
        cancel = main[main.index("Public Sub CancelPriceLookup(") :]
        self.assertIn("Application.EnableEvents = False", cancel[: cancel.index("End Sub")])  # no refresh on close
        self.assertIn("M_Ribbon.RefreshRibbon", cancel[: cancel.index("End Sub")])  # a kept workbook keeps its Ribbon
        self.assertIn('"Downloading the prices of "', main)  # the wait shows in the status bar
        self.assertNotIn("from STOCKHISTORY?", main)  # the confirmation does not name the function
        self.assertIn('CallByName PriceCell, "Formula2", VbLet', main)  # late bound: it compiles in any Excel
        self.assertRegex(main, r"(?m)^Public Sub DownloadClosingPrices\(\)")
        self.assertRegex(main, r"(?m)^Public Function DownloadClosingPrices_Silent\( _\n    Optional ByVal Source As String = vbNullString\) _\n    As String")
        for removed in ("CopyClosingPrices", "ScanClosingPrices", "FX_COL_PRICE_PAIR"):
            self.assertNotIn(removed, main)

    def test_the_fx_table_is_filled_with_twenty_currencies_against_the_euro(self) -> None:
        first = self.row("FX_Table") + 1
        rows = range(first, first + FX_ROWS)
        codes = [self.text(f"C{row}") for row in rows]
        self.assertEqual(codes[0], "EUR")
        self.assertEqual(len(set(codes)), FX_ROWS)
        self.assertTrue(all(re.fullmatch(r"[A-Z]{3}", code) for code in codes), codes)
        for row in rows:
            self.assertTrue(self.text(f"D{row}"), row)
            self.assertGreater(float(self.text(f"E{row}")), 0, row)
        self.assertEqual(self.text(f"E{first}"), "1")  # the reporting currency against itself
        self.assertEqual(len({self.text(f"E{row}") for row in rows[1:]}), FX_ROWS - 1)  # no copied placeholder
        below = first + FX_ROWS
        self.assertTrue(self.text(f"D{below}").startswith("Add currencies directly below the last row"))
        self.assertEqual(self.text(f"C{below}"), "")  # a note in column C would be read as a currency

    def test_the_parameters_are_in_the_owners_order_with_their_names(self) -> None:
        # Each parameter is a row: its code in column C, its value in column E under its own workbook name.
        groups = (["AsOfDate", "ReportingCcy", "Regime", "Alpha", "MultiplierFloor"],
                  ["LambdaThresholdIR", "LambdaThresholdCO", "DaycountConvention", "YearConvention", "Calendar"])
        for group in groups:
            rows = [self.row(name) for name in group]
            self.assertEqual(rows, list(range(rows[0], rows[0] + len(group))), group)  # consecutive, in this order
            for name, row in zip(group, rows):
                self.assertEqual(self.at(name).column, "E", name)
                self.assertEqual(self.text(f"C{row}"), name)
        self.assertLess(self.row("MultiplierFloor"), self.row("LambdaThresholdIR"))
        self.assertEqual(self.row("AsOfDate"), self.row("Params_Code") + 1)  # the first parameter under the header
        for old in ("DayCountConvention", "DayCountNumerator"):
            self.assertNotIn(f'name="{old}"', self.workbook)
        # a validation belongs to the cell it checks: the regime list on Regime, the currency check on ReportingCcy only
        regime, currency = self.row("Regime"), self.row("ReportingCcy")
        self.assertEqual(self.validation(f"E{regime}"), '"CRR,BCBS"')
        self.assertRegex(self.xml, rf'<dataValidation [^>]*sqref="E{currency}"')
        self.assertNotRegex(self.xml, rf'sqref="E{currency}:E{regime}"')
        self.assertEqual(self.validation(f"E{self.row('DaycountConvention')}"), '"ACTUAL"')

    def test_the_rate_header_is_ten_points_as_the_style_rules_allow(self) -> None:
        styles = ElementTree.fromstring(self.package.read("xl/styles.xml"))
        xfs = list(styles.find(MAIN + "cellXfs"))  # type: ignore[arg-type]
        fonts = list(styles.find(MAIN + "fonts"))  # type: ignore[arg-type]
        cell = self.sheet.find(f'.//{MAIN}c[@r="E{self.row("FX_Table")}"]')
        assert cell is not None
        font = fonts[int(xfs[int(cell.attrib["s"])].attrib["fontId"])]
        size = font.find(MAIN + "sz")
        assert size is not None
        self.assertEqual(size.attrib["val"], "10")
        rules = (ROOT / "docs" / "workbook" / "sheet-style.json").read_text(encoding="utf-8")
        self.assertIn('"header_size_exceptions"', rules)
        self.assertIn('"Params": {"FX_Table:E": 10}', rules)

    def test_the_date_of_the_prices_is_shown_beside_the_button(self) -> None:
        date, stamp = self.at("FX_PriceDate"), f"G{self.row('FX_Table') - 1}"  # the stamp: on the row of the button
        self.assertEqual(date.column, "AQ")
        # the date of the prices in the template: the reporting date
        self.assertEqual(self.text(f"AQ{date.row}"), self.text(f"E{self.row('AsOfDate')}"))
        self.assertEqual(self.text(stamp), "Closing prices of 2026-09-30")
        formula = self.sheet.find(f'.//{MAIN}c[@r="{stamp}"]/{MAIN}f')
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

    def test_the_source_of_the_rates_is_chosen_beside_the_button_and_stated_with_their_date(self) -> None:
        # #745: FX_Source, on the row above the FX table, chooses what Download prices writes; FX_PriceSource
        # keeps the source of the rates downloaded, which the stamp states with their date, so that the rates in
        # the table never have an unknown source.
        source, kept = self.at("FX_Source"), self.at("FX_PriceSource")
        self.assertEqual((source.column, source.row), ("D", self.row("FX_Table") - 1))
        self.assertEqual((kept.column, kept.row), ("AQ", self.row("FX_PriceDate") + 1))
        self.assertEqual(self.text(f"D{source.row}"), "Closing price")
        self.assertEqual(self.text(f"AQ{kept.row}"), "Closing price")
        self.assertEqual(self.validation(f"D{source.row}"), '"Closing price,ECB reference rate"')
        formula = self.sheet.find(f'.//{MAIN}c[@r="G{source.row}"]/{MAIN}f')
        assert formula is not None and formula.text is not None
        self.assertIn('IF(FX_PriceSource="ECB reference rate","ECB reference rates of ","Closing prices of ")', formula.text)
        config = (ROOT / "src" / "core" / "CORE_Config.bas").read_text(encoding="cp1252")
        for constant in ('NM_FX_SOURCE     As String = "FX_Source"', 'NM_FX_PRICE_SRC  As String = "FX_PriceSource"',
                         'FX_SOURCE_CLOSING  As String = "Closing price"', 'FX_SOURCE_ECB      As String = "ECB reference rate"'):
            self.assertIn(constant, config)

    def test_the_ecb_is_asked_over_https_with_currency_codes_and_dates_only(self) -> None:
        # #745: one HTTPS request from VBA, no web query and no connection left in the workbook, nothing sent
        # but three-letter codes and two dates, no credential.
        config = (ROOT / "src" / "core" / "CORE_Config.bas").read_text(encoding="cp1252")
        self.assertIn('ECB_EXR_URL        As String = "https://data-api.ecb.europa.eu/service/data/EXR/D."', config)
        main = (ROOT / "src" / "workbook" / "M_Main.bas").read_text(encoding="cp1252")
        url = main[main.index("Public Function EcbRequestUrl("):]
        url = url[: url.index("End Function")]
        self.assertIn('ECB_EXR_URL & Join(Codes, "+") & ".EUR.SP00.A?startPeriod="', url)
        self.assertIn('"&format=csvdata&detail=dataonly"', url)
        self.assertIn('Codes(i) Like "[A-Z][A-Z][A-Z]"', main)  # only codes go into the request
        self.assertIn('CreateObject("MSXML2.XMLHTTP.6.0")', main)
        self.assertIn('LCase$(Left$(Url, 8)) <> "https://"', main)
        answer = main[main.index("Private Function EcbAnswer("):]
        self.assertIn('Case "ecb-answer", "ecb-write"', answer)  # no test fault sends a real request
        # the row of the reporting currency takes 1, and only while it is still that currency's row
        self.assertEqual(main.count("If Not IsReportingRow("), 2)
        for absent in ("QueryTables", "Connections.Add", "WorkbookConnection", '"Authorization"', "SetCredentials"):
            self.assertNotIn(absent, main)

    def test_the_sample_answer_of_the_tests_is_the_fixture_file(self) -> None:
        tests = (ROOT / "tests" / "modules" / "TEST_MainState.bas").read_text(encoding="cp1252")
        block = re.search(r"Private Function EcbSample\(\).*?Join\(Array\((.*?)\), vbLf\)", tests, re.S)
        assert block is not None
        lines = [item.replace('""', '"') for item in re.findall(r'"((?:[^"]|"")*)"', block.group(1))]
        fixture = (ROOT / "tests" / "fixtures" / "ecb" / "EXR_D_sample.csv").read_text(encoding="utf-8").splitlines()
        self.assertEqual(lines, fixture)
        self.assertEqual(fixture[0], "KEY,FREQ,CURRENCY,CURRENCY_DENOM,EXR_TYPE,EXR_SUFFIX,TIME_PERIOD,OBS_VALUE")

    def test_a_change_of_the_reporting_currency_offers_to_rebase_the_fx_table(self) -> None:
        # #725: the rates are relative to the reporting currency, so a new one is offered a rebase; never in
        # silent mode or during an operation, so that tests and automation are never stopped by a question.
        workbook = (ROOT / "src" / "workbook" / "ThisWorkbook.cls").read_text(encoding="cp1252")
        change = workbook[workbook.index("Private Sub Workbook_SheetChange"):]
        self.assertIn("M_Main.ReportingCurrencyChanged Sh, Target", change[: change.index("End Sub")])
        main = (ROOT / "src" / "workbook" / "M_Main.bas").read_text(encoding="cp1252")
        handler = main[main.index("Public Sub ReportingCurrencyChanged("):]
        handler = handler[: handler.index("\nEnd Sub")]
        self.assertIn("If mRunning Or gSilent Then", handler)
        self.assertIn("vbYesNo + vbQuestion + vbDefaultButton2", handler)
        self.assertIn("RebaseFxTable(Code, Problem)", handler)
        engine = (ROOT / "src" / "core" / "CORE_Engine.bas").read_text(encoding="cp1252")
        self.assertIn("accept the rebase of the FX table the workbook offers, press", engine)

    def test_a_currency_appended_below_the_table_has_a_pair_from_its_code(self) -> None:
        main = (ROOT / "src" / "workbook" / "M_Main.bas").read_text(encoding="cp1252")
        # the pair is built from the code and the reporting currency in code, for every row down to the first blank key
        self.assertIn("ClosingPair = Reporting & \"/\" & Code", main)

    def test_the_closing_day_validation_and_the_unlocked_cells_follow_the_table(self) -> None:
        first = self.row("Holiday_Table") + 1
        # 61 is a date serial, 1900-03-01, not a row: a closing day is a real date
        self.assertRegex(self.xml, rf'sqref="I{first}:I{first + HOLIDAY_ROWS - 1}"[^>]*>'
                                   rf'<formula1>AND\(ISNUMBER\(I{first}\),I{first}&gt;=61\)</formula1>')
        self.assertNotRegex(self.xml, r'sqref="G\d+:G\d+"[^>]*><formula1>AND\(ISNUMBER\(G')
        matrix = (ROOT / "docs" / "workbook" / "protection.json").read_text(encoding="utf-8")
        self.assertIn(f'"FX_Table+{FX_SPARE}", "Holiday_Table+{HOLIDAY_ROWS}"', matrix)
        self.assertIn('"unlocked_cells": ["AsOfDate", "ReportingCcy", "Regime", "YearConvention", "Calendar", "FX_Source"',
                      matrix)

    def test_the_trades_validations_read_the_moved_tables(self) -> None:
        trades = self.package.read(sheet_parts(self.package)["Trades"]).decode("utf-8")
        factors, currencies = self.row("SF_Table") + 1, self.row("FX_Table") + 1
        self.assertIn(f"Params!$E${factors}:$E${factors + 25}", trades)
        self.assertEqual(sorted(set(re.findall(r"Params!\$C\$\d+:\$C\$\d+", trades))),
                         [f"Params!$C${currencies}:$C${currencies + FX_ROWS - 1}"])

    def test_the_fx_table_holds_the_euro_and_twenty_currencies(self) -> None:
        # the reporting currency and 20 others, validated and checked for a repeated code
        first = self.row("FX_Table") + 1
        last = first + FX_ROWS - 1
        self.assertRegex(self.xml, rf'sqref="C{first}:C{last}"[^>]*><formula1>AND\(ISTEXT\(C{first}\),LEN\(C{first}\)=3,'
                                   rf'COUNTIF\(\$C\${first}:\$C\${last},C{first}\)=1\)')
        self.assertRegex(self.xml, rf'<dataValidation [^>]*sqref="E{first}:E{last}"')
        self.assertRegex(self.xml, rf'<conditionalFormatting sqref="C{first}:C{last}">')
        self.assertNotIn('ReportingCcy&amp;"/"&amp;C', self.xml)  # no pair formula: the macro builds the pair


if __name__ == "__main__":
    unittest.main()
