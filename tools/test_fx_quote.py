"""Tests of the FX table convention of the workbook template and the engine (#725).

The rate column is the units of the currency per 1 unit of the reporting currency, as a market quotes it.
"""

from __future__ import annotations

import re
import unittest
import zipfile
from datetime import date
from math import exp
from pathlib import Path
from xml.etree import ElementTree

from test_template_buttons import sheet_parts

ROOT = Path(__file__).resolve().parents[1]
TEMPLATE = ROOT / "src" / "workbook" / "SACCR_Template.xlsx"
MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


class FxQuoteTests(unittest.TestCase):
    def setUp(self) -> None:
        self.package = zipfile.ZipFile(TEMPLATE)
        self.addCleanup(self.package.close)
        self.config = (ROOT / "src" / "core" / "CORE_Config.bas").read_text(encoding="cp1252")

    def headers(self) -> tuple[str, str]:
        header = re.search(r'FX_HEADERS As String = "Currency Code\|\|([^"]+)"', self.config)
        before = re.search(r'FX_HEADER_BEFORE_725 As String = "([^"]+)"', self.config)
        self.assertIsNotNone(header)
        self.assertIsNotNone(before)
        assert header is not None and before is not None
        return header.group(1), before.group(1)

    def test_the_header_of_the_rate_column_changed_and_the_template_has_the_new_one(self) -> None:
        header, before = self.headers()
        self.assertNotEqual(header, before)
        strings = self.package.read("xl/sharedStrings.xml").decode("utf-8")
        self.assertIn(f"<t>{header}</t>", strings)
        self.assertNotIn(f"<t>{before}</t>", strings)

    def test_the_illustrative_rates_are_market_style(self) -> None:
        parts = sheet_parts(self.package)
        sheet = ElementTree.fromstring(self.package.read(parts["Params"]))
        rates = {item.attrib["r"]: float(item.findtext(f"{MAIN}v") or "nan")
                 for item in sheet.iter(f"{MAIN}c") if item.attrib["r"] in {"E61", "E62", "E63", "E64"}}
        self.assertEqual(rates, {"E61": 1.0, "E62": 1.1328, "E63": 178.29, "E64": 0.8536})

    def test_the_engine_divides_by_the_rate(self) -> None:
        engine = (ROOT / "src" / "core" / "CORE_Engine.bas").read_text(encoding="cp1252")
        self.assertIn("Trade.notionalRep = Trade.notional / Trade.fxN", engine)
        self.assertIn("Trade.mtmRep = Trade.mtm / Trade.fxM", engine)
        self.assertNotIn("Trade.notional * Trade.fxN", engine)

    def test_an_old_header_is_refused_with_the_conversion_note(self) -> None:
        engine = (ROOT / "src" / "core" / "CORE_Engine.bas").read_text(encoding="cp1252")
        self.assertIn("FX_HEADER_BEFORE_725", engine)
        self.assertIn("invert every rate and rename the header", engine)

    def test_the_multi_currency_cases_agree_with_a_separate_derivation(self) -> None:
        # #725: the expected values of the cases fx-rate-old-convention-value and fx-rate-multi-currency of
        # TEST_InputValidation, derived here from the rules and the market quote, sharing no code with the engine.
        # Five-year swaps from the reporting date, S = 0: supervisory duration CRE52.34, CRR Art. 279b(1)(a).
        years = (date(2031, 9, 30) - date(2026, 9, 30)).days / 365.0
        duration = (1.0 - exp(-0.05 * years)) / 0.05
        usd, jpy, gbp = 1.25, 160.0, 0.85  # units of the currency per 1 EUR
        old_value = 10000.0 * duration / 0.8  # 0.80, the old convention's rate, read as USD per EUR
        usd_swap = 10000.0 / usd * duration
        eur_swap = 20000.0 * duration
        forward = 1600000.0 / jpy  # FX adjusted notional: the JPY leg in EUR, CRE52.35, CRR Art. 279b(1)(b)
        value = 500.0 / usd - 300.0 / gbp  # MtM in USD and in GBP, each divided by its rate
        replacement_cost = max(value, 0.0)  # unmargined, no collateral: CRE52.10, CRR Art. 275(1)
        ir_add_on = 0.005 * abs(usd_swap) + 0.005 * abs(eur_swap)  # one hedging set per currency, CRE52.57
        fx_add_on = 0.04 * abs(forward)  # CRE52.58, CRR Art. 280b
        exposure = 1.4 * (replacement_cost + ir_add_on + fx_add_on)  # V - C > 0: multiplier 1
        expected = {"55326.4736651278": old_value, "35408.9431456818": usd_swap, "10000": forward,
                    "47.0588235294118": replacement_cost, "619.656505049432": ir_add_on, "400": fx_add_on,
                    "1493.40146001038": exposure}
        tests = (ROOT / "tests" / "modules" / "TEST_InputValidation.bas").read_text(encoding="cp1252")
        start = tests.index('StartCase "fx-rate-old-convention-value"')
        block = tests[start:tests.index("StartCase", tests.index('StartCase "fx-rate-multi-currency"') + 1)]
        for text, derived in expected.items():
            self.assertIn(f'"{text}"', block, text)
            self.assertAlmostEqual(float(text), derived, delta=1e-9 * max(1.0, abs(derived)), msg=text)
        self.assertNotAlmostEqual(old_value, usd_swap, delta=1.0)  # the old value is a different input


if __name__ == "__main__":
    unittest.main()
