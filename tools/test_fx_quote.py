"""Tests of the FX table convention of the workbook template and the engine (#725).

The rate column is the units of the currency per 1 unit of the reporting currency, as a market quotes it.
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


if __name__ == "__main__":
    unittest.main()
