"""Tests of the CSV import specification page and its example records (#473)."""
import copy
import unittest
from html.parser import HTMLParser
from pathlib import Path
from typing import Any

import check_csv_contract as gate
import import_spec as spec
from csv_contract import load_contract, validate

ROOT = Path(__file__).resolve().parents[1]


class PreText(HTMLParser):
    """The text of each <pre> element by its id, as a browser copies it."""

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.texts: dict[str, str] = {}
        self.current: str | None = None

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag == "pre":
            self.current = dict(attrs).get("id") or ""
            self.texts[self.current] = ""

    def handle_endtag(self, tag: str) -> None:
        if tag == "pre":
            self.current = None

    def handle_data(self, data: str) -> None:
        if self.current is not None:
            self.texts[self.current] += data


class ImportSpecificationTests(unittest.TestCase):
    contract: dict[str, Any]
    context: dict[str, Any]
    page: str

    @classmethod
    def setUpClass(cls) -> None:
        cls.contract = load_contract(ROOT)
        cls.context = gate.template_context(gate.Workbook(ROOT / gate.TEMPLATE))
        cls.page = spec.render(cls.contract)

    def test_committed_page_is_current(self) -> None:
        self.assertEqual(spec.specification_findings(ROOT, self.contract), [])

    def test_template_context(self) -> None:
        self.assertEqual(self.context["as_of"], "2026-09-30")
        self.assertEqual(self.context["reporting_currency"], "EUR")
        self.assertEqual(self.context["regime"], "CRR")
        self.assertIn("USD", self.context["currencies"])
        self.assertEqual(self.context["factors"]["OT"], "CRR")
        self.assertEqual(self.context["factors"]["IR"], "")

    def test_examples_import_together(self) -> None:
        self.assertEqual(spec.example_findings(self.contract, self.context), [])

    def test_copied_examples_are_the_files(self) -> None:
        parser = PreText()
        parser.feed(self.page)
        for kind, anchor in (("NettingSets", "example-nettingsets"), ("Trades", "example-trades")):
            with self.subTest(kind=kind):
                copied = parser.texts[anchor]
                self.assertEqual(copied, spec.example_text(kind, self.contract))
                rows, findings = validate(copied.encode("utf-8"), kind, self.contract)
                self.assertEqual(findings, [])
                names = [column["name"] for column in self.contract["files"][kind]["columns"]]
                self.assertEqual(rows, [{name: record.get(name, "") for name in names}
                                        for record in self.contract["files"][kind]["examples"]])

    def test_examples_cover_the_cases_the_page_explains(self) -> None:
        netting_sets = spec.example_text("NettingSets", self.contract)
        self.assertIn('"Müller & Søn ""Nord""; <Example> AG"', netting_sets)
        self.assertIn("\n00731;", netting_sets)
        trades = [row for _, row in spec.example_rows("Trades", self.contract)]
        self.assertTrue(any(row["trade_id"].startswith("0") for row in trades))
        self.assertTrue(any(row["economic_trade_id"] and row["mtm_leg"] == "N" for row in trades))
        self.assertTrue(any(row["settlement_type"] == "STM" for row in trades))
        self.assertTrue(any(not row["notional_currency"] or not row["comment"] for row in trades))

    def test_text_is_escaped_and_inert(self) -> None:
        self.assertIn("&lt;Example&gt;", self.page)
        self.assertNotIn("<Example>", self.page)
        lowered = self.page.lower()
        for fragment in ("<script", "javascript:", "http://", "https://", "src=", "@import", "url(", "<link"):
            self.assertNotIn(fragment, lowered)

    def test_page_states_the_contract(self) -> None:
        self.assertIn(f"contract version {self.contract['contract_version']}", self.page)
        for kind, file_spec in self.contract["files"].items():
            self.assertIn(f"{file_spec['max_records']:,}", self.page)
            for column in file_spec["columns"]:
                self.assertIn(f"<code>{column['name']}</code>", self.page, kind)
        for code in self.contract["errors"]:
            self.assertIn(f"<code>{code}</code>", self.page)
        self.assertIn(f"<code>{self.contract['format']['delimiter']}</code>", self.page)

    def test_quoting(self) -> None:
        self.assertEqual(spec.csv_field("plain"), "plain")
        self.assertEqual(spec.csv_field("a, b"), "a, b")
        self.assertEqual(spec.csv_field("a;b"), '"a;b"')
        self.assertEqual(spec.csv_field('say "x"'), '"say ""x"""')
        self.assertEqual(spec.csv_field("two\nlines"), '"two\nlines"')

    def test_example_count_is_bounded(self) -> None:
        contract = copy.deepcopy(self.contract)
        contract["files"]["NettingSets"]["examples"] = contract["files"]["NettingSets"]["examples"][:2]
        self.assertEqual(spec.example_findings(contract, self.context),
                         ["NettingSets: 2 example records, the specification shows 3 to 5"])

    def test_broken_examples_are_found(self) -> None:
        contract = copy.deepcopy(self.contract)
        contract["files"]["Trades"]["examples"][0]["netting_set_id"] = "NS-404"
        findings = spec.example_findings(contract, self.context)
        self.assertTrue(any("CSV-NETTING-SET-UNKNOWN" in finding for finding in findings), findings)
        contract = copy.deepcopy(self.contract)
        context = dict(self.context, as_of="2029-11-01")
        findings = spec.example_findings(contract, context)
        self.assertTrue(any("CSV-CONTEXT-MATURED" in finding for finding in findings), findings)
        contract["files"]["Trades"]["examples"][0]["notional_amount"] = "1,5"
        findings = spec.example_findings(contract, self.context)
        self.assertTrue(any("CSV-TYPE" in finding for finding in findings), findings)

    def test_unknown_example_column(self) -> None:
        contract = copy.deepcopy(self.contract)
        contract["files"]["Trades"]["examples"][0]["notional"] = "1"
        self.assertIn("Trades: example 1 names unknown column(s) ['notional']",
                      spec.example_findings(contract, self.context))


if __name__ == "__main__":
    unittest.main()
