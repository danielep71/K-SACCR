"""Deterministic tests of the CSV input contract and its reference implementation (#206)."""
import json
import shutil
import tempfile
import unittest
import zipfile
from pathlib import Path
from typing import Any

import check_csv_contract as gate
import csv_contract as csv

ROOT = Path(__file__).resolve().parents[1]
CONTRACT: dict[str, Any] = csv.load_contract(ROOT)
NS_HEADER = csv.template_text("NettingSets", CONTRACT).rstrip("\n")
TR_HEADER = csv.template_text("Trades", CONTRACT).rstrip("\n")
NS_ROW = "NS-1;CPTY;Y;N;1;N;N;;0;0;0;0;;;;;"
TR_ROW = "T1;NS-1;IR;;EUR;Linear;Long;;;;1000;EUR;0;;;2031-01-02;2031-01-02;;;;;;;;;;;"


def codes(data: str | bytes, kind: str = "NettingSets", **options: Any) -> list[tuple[str, int, int, str]]:
    raw = data.encode("utf-8") if isinstance(data, str) else data
    return [(f.code, f.line, f.record, f.column) for f in csv.validate(raw, kind, CONTRACT, **options)[1]]


class ParserTests(unittest.TestCase):
    def test_quoting_rules(self) -> None:
        records, findings = csv.parse('a;"b;c";"d ""e""";\r\n"multi\nline";x;;\n')
        self.assertEqual(findings, [])
        self.assertEqual([r.fields for r in records], [["a", "b;c", 'd "e"', ""], ["multi\nline", "x", "", ""]])
        self.assertEqual([r.line for r in records], [1, 2])

    def test_a_comma_does_not_separate_fields(self) -> None:
        records, findings = csv.parse("1,5;a,b\n")
        self.assertEqual((findings, [r.fields for r in records]), ([], [["1,5", "a,b"]]))

    def test_line_numbers_follow_quoted_line_breaks(self) -> None:
        records, _ = csv.parse('h\n"one\ntwo\nthree"\nnext\n')
        self.assertEqual([r.line for r in records], [1, 2, 5])

    def test_syntax_errors(self) -> None:
        for text, message in (('a"b\n', "quote inside"), ('"a"b\n', "after a closing quote"),
                              ('"open\n', "not closed"), ("a\rb\n", "carriage return")):
            with self.subTest(text=text):
                _, findings = csv.parse(text)
                self.assertEqual([f.code for f in findings], ["CSV-SYNTAX"])
                self.assertIn(message, findings[0].message)

    def test_last_record_needs_no_line_end(self) -> None:
        records, findings = csv.parse("a;b\nc;d")
        self.assertEqual((findings, [r.fields for r in records]), ([], [["a", "b"], ["c", "d"]]))


class EncodingTests(unittest.TestCase):
    def test_bom_and_crlf_variants_of_every_fixture_match_lf(self) -> None:
        for path in sorted((ROOT / gate.FIXTURES).glob("*/*.csv")):
            kind = "Trades" if path.name.startswith("trades") else "NettingSets"
            plain = path.read_bytes()
            with self.subTest(path=path.name):
                expected = codes(plain, kind)
                self.assertEqual(codes(csv.BOM + plain, kind), expected)
                self.assertEqual(codes(plain.replace(b"\n", b"\r\n"), kind), expected)

    def test_invalid_utf8_and_nul(self) -> None:
        self.assertEqual(codes(NS_HEADER.encode() + b"\n\xff\n"), [("CSV-ENCODING", 0, 0, "")])
        self.assertEqual(codes(NS_HEADER + "\n\x00\n")[0][0], "CSV-ENCODING")

    def test_non_ascii_text_is_kept(self) -> None:
        rows, findings = csv.validate((NS_HEADER + "\n" + NS_ROW.replace("CPTY", "Société Générale")).encode(),
                                      "NettingSets", CONTRACT)
        self.assertEqual((findings, rows[0]["counterparty"]), ([], "Société Générale"))


class FileTests(unittest.TestCase):
    def test_empty_file_and_header_only(self) -> None:
        self.assertEqual(codes(b""), [("CSV-EMPTY", 1, 0, "")])
        self.assertEqual(codes(csv.BOM), [("CSV-EMPTY", 1, 0, "")])
        self.assertEqual(csv.validate((NS_HEADER + "\n").encode(), "NettingSets", CONTRACT), ([], []))

    def test_templates_round_trip(self) -> None:
        for kind in CONTRACT["files"]:
            with self.subTest(kind=kind):
                self.assertEqual(csv.validate(csv.template_text(kind, CONTRACT).encode(), kind, CONTRACT), ([], []))

    def test_a_header_error_stops_reading(self) -> None:
        bad_header = NS_HEADER.replace("netting_set_id", "bad_header")
        self.assertEqual(codes(bad_header + '\n"unclosed,quote\n'),
                         [("CSV-HEADER-UNKNOWN", 1, 0, "bad_header"), ("CSV-HEADER-MISSING", 1, 0, "netting_set_id")])
        self.assertEqual(codes(NS_HEADER + '\n"unclosed,quote\n'), [("CSV-SYNTAX", 2, 1, "")])

    def test_another_delimiter_is_one_finding(self) -> None:
        for other in (",", "\t"):
            with self.subTest(delimiter=other):
                text = NS_HEADER.replace(";", other) + "\n" + NS_ROW.replace(";", other) + "\n"
                self.assertEqual(codes(text), [("CSV-DELIMITER", 1, 0, "")])
        self.assertEqual(codes(NS_HEADER.replace(";", ",", 1) + "\n")[0], ("CSV-HEADER-UNKNOWN", 1, 0, "netting_set_id,counterparty"))

    def test_header_order_is_free_and_names_are_exact(self) -> None:
        names = NS_HEADER.split(";")
        values = NS_ROW.split(";")
        reordered = ";".join(reversed(names)) + "\n" + ";".join(reversed(values)) + "\n"
        rows, findings = csv.validate(reordered.encode(), "NettingSets", CONTRACT)
        self.assertEqual((findings, rows[0]["netting_set_id"]), ([], "NS-1"))
        self.assertEqual(codes(NS_HEADER.replace("margined", "Margined") + "\n"),
                         [("CSV-HEADER-UNKNOWN", 1, 0, "Margined"), ("CSV-HEADER-MISSING", 1, 0, "margined")])

    def test_record_limit_rejects_the_whole_file(self) -> None:
        text = NS_HEADER + "\n" + "\n".join(NS_ROW.replace("NS-1", f"NS-{i}") for i in range(3)) + "\n"
        self.assertEqual(codes(text, max_records=2), [("CSV-TOO-MANY-RECORDS", 4, 3, "")])
        self.assertEqual(codes(text, max_records=3), [])

    def test_a_rejected_file_returns_no_records(self) -> None:
        text = NS_HEADER + "\n" + NS_ROW + "\n" + NS_ROW.replace("NS-1", "NS-2").replace(";Y;N;", ";X;N;") + "\n"
        rows, findings = csv.validate(text.encode(), "NettingSets", CONTRACT)
        self.assertEqual((rows, [f.code for f in findings]), ([], ["CSV-ALLOWED"]))


class FieldTests(unittest.TestCase):
    def check_trade(self, column: str, value: str) -> list[str]:
        names = TR_HEADER.split(";")
        values = TR_ROW.split(";")
        values[names.index(column)] = value
        return [code for code, _, _, _ in codes(TR_HEADER + "\n" + ";".join(values) + "\n", "Trades")]

    def test_numbers(self) -> None:
        for value in ("1000", "0.5", "1.25e6", "1E+7", "123456789012345", "1000000000000000000"):
            with self.subTest(value=value):
                self.assertEqual(self.check_trade("notional_amount", value), [])
        for value in ("+5", "1,000", "1,5", '"1;000"', "1 000", "0x10", ".5", "5.", "007", "1234567890123456", "NaN", "inf"):
            with self.subTest(value=value):
                self.assertEqual(self.check_trade("notional_amount", value), ["CSV-TYPE"])
        self.assertEqual(self.check_trade("market_value_amount", "-250000.75"), [])
        self.assertEqual(self.check_trade("notional_amount", "0"), ["CSV-RANGE"])
        for value in ("1e9999", "-1e309", "1e-9999", "1e308"):
            with self.subTest(value=value):
                self.assertIn(self.check_trade("market_value_amount", value), (["CSV-RANGE"],))
        self.assertEqual(self.check_trade("market_value_amount", "0e-9999"), [])

    def test_dates(self) -> None:
        self.assertEqual(self.check_trade("maturity_date", "2200-12-31"), [])
        for value in ("2031-1-2", "02/01/2031", "2031-02-29", "2031-01-02T00:00", "47756"):
            with self.subTest(value=value):
                self.assertEqual(self.check_trade("maturity_date", value), ["CSV-TYPE"])
        self.assertEqual(self.check_trade("maturity_date", "2201-01-01"), ["CSV-RANGE"])

    def test_identifiers_keep_leading_zeros_and_reject_formulas(self) -> None:
        self.assertEqual(self.check_trade("trade_id", "000123"), [])
        self.assertEqual(self.check_trade("trade_id", "=1+1"), ["CSV-FORMULA"])
        self.assertEqual(self.check_trade("trade_id", "-123"), ["CSV-FORMULA"])
        self.assertEqual(self.check_trade("trade_id", "A" * 65), ["CSV-LENGTH"])
        self.assertEqual(self.check_trade("trade_id", "T#1"), ["CSV-TYPE"])
        self.assertEqual(self.check_trade("comment", '"@SUM(A1)"'), ["CSV-FORMULA"])
        self.assertEqual(self.check_trade("comment", '"a\nb"'), [])
        for char in ("\x7f", "\x85", "\x9f"):
            with self.subTest(char=hex(ord(char))):
                self.assertEqual(self.check_trade("trade_id", "T" + char), ["CSV-TYPE"])
                self.assertEqual(self.check_trade("comment", "note" + char), ["CSV-TYPE"])

    def test_codes_are_exact(self) -> None:
        self.assertEqual(self.check_trade("instrument", "Option"), [])
        self.assertEqual(self.check_trade("instrument", "OPTION"), ["CSV-ALLOWED"])
        self.assertEqual(self.check_trade("mtm_leg", "y"), ["CSV-ALLOWED"])
        self.assertEqual(self.check_trade("notional_currency", "Eur"), ["CSV-TYPE"])


def netting_set(role: str, qccp: str = "", retains: str = "") -> str:
    values = NS_ROW.split(";")
    values[3], values[-2], values[-1] = role, qccp, retains
    return NS_HEADER + "\n" + ";".join(values) + "\n"


class RowRuleTests(unittest.TestCase):
    def test_a_rule_is_not_judged_on_a_flawed_predicate(self) -> None:
        self.assertEqual(codes(netting_set("BAD", "Y", "N")), [("CSV-ALLOWED", 2, 1, "clearing_role")])

    def test_a_rule_is_not_judged_on_a_flawed_target(self) -> None:
        self.assertEqual(codes(netting_set("CCP", "BAD")), [("CSV-ALLOWED", 2, 1, "qccp")])

    def test_conditional_rule_both_ways(self) -> None:
        self.assertEqual([c[3] for c in codes(netting_set("CCP"))], ["ccp_retains_variation_margin", "qccp"])
        self.assertEqual(codes(netting_set("CCP", "Y", "N")), [])
        self.assertEqual([c[3] for c in codes(netting_set("CM", "Y"))], ["qccp"])


class GateTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        for path in (csv.CONTRACT, csv.DOCUMENT, gate.CONFIG, gate.TEMPLATE, gate.EXPECTED, gate.CASES):
            (self.root / path).parent.mkdir(parents=True, exist_ok=True)
            shutil.copy(ROOT / path, self.root / path)
        shutil.copytree(ROOT / csv.TEMPLATES, self.root / csv.TEMPLATES)
        for folder in ("valid", "invalid"):
            shutil.copytree(ROOT / gate.FIXTURES / folder, self.root / gate.FIXTURES / folder)

    def edit_contract(self, change: Any) -> None:
        path = self.root / csv.CONTRACT
        contract = json.loads(path.read_text(encoding="utf-8"))
        change(contract)
        path.write_text(json.dumps(contract), encoding="utf-8")

    def test_repository_passes(self) -> None:
        self.assertEqual(gate.run_check(self.root)["findings"], [])

    def test_allowed_values_must_match_the_template(self) -> None:
        self.edit_contract(lambda c: c["files"]["Trades"]["columns"][5]["allowed"].append("Swap"))
        self.assertTrue(any("instrument: allowed values" in f for f in gate.run_check(self.root)["findings"]))

    def test_sheet_headers_must_match_the_engine(self) -> None:
        self.edit_contract(lambda c: c["files"]["NettingSets"]["columns"][2].update(sheet_header="Margin flag"))
        findings = gate.run_check(self.root)["findings"]
        self.assertIn(f"NettingSets: sheet headers differ from NS_HEADERS in {gate.CONFIG}", findings)

    def test_templates_must_be_regenerated(self) -> None:
        self.edit_contract(lambda c: c["files"]["NettingSets"]["columns"][1].update(name="counterparty_name"))
        self.assertTrue(any("is not what the contract generates" in f for f in gate.run_check(self.root)["findings"]))

    def test_export_constants_must_match_the_contract(self) -> None:
        config = self.root / gate.CONFIG
        text = config.read_bytes()
        config.write_bytes(text.replace(b'"K-SACCR_Trades_v1.csv"', b'"K-SACCR_Trades.csv"')
                           .replace(b'"trade_id;netting_set_id;', b'"netting_set_id;trade_id;'))
        findings = gate.run_check(self.root)["findings"]
        self.assertIn(f"Trades: CSV_TR_FILE in {gate.CONFIG} differs from the template name K-SACCR_Trades_v1.csv",
                      findings)
        self.assertIn(f"Trades: CSV_TR_HEADER in {gate.CONFIG} differs from the contract's header line", findings)
        self.assertFalse(any(f.startswith("NettingSets: CSV_") for f in findings))

    def test_delimiter_must_be_the_parsers(self) -> None:
        self.edit_contract(lambda c: c["format"].update(delimiter=","))
        self.assertIn("format: delimiter ',' is not the parser's ';'", gate.run_check(self.root)["findings"])

    def edit_template(self, part: str, change: Any) -> None:
        path = self.root / gate.TEMPLATE
        with zipfile.ZipFile(path) as package:
            parts = {item: package.read(item) for item in package.namelist()}
        parts[part] = change(parts[part].decode("utf-8")).encode("utf-8")
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as package:
            for item, data in parts.items():
                package.writestr(item, data)

    def test_staging_sheet_must_follow_the_contract(self) -> None:
        self.edit_template("xl/workbook.xml", lambda x: x.replace("Import_Trades!$C$9", "Import_Trades!$C$8"))
        self.edit_contract(lambda c: c["files"]["NettingSets"]["columns"].reverse())
        findings = gate.run_check(self.root)["findings"]
        self.assertIn("Trades: ImportTrades_Header is 'Import_Trades!$C$8', not Import_Trades!$C$9", findings)
        self.assertIn("NettingSets: the Import_NettingSets header row differs from the contract's sheet headers",
                      findings)
        self.assertIn("NettingSets: the CSV names above the Import_NettingSets header row differ from the "
                      "contract's columns", findings)

    def test_fixture_outcomes_are_pinned(self) -> None:
        (self.root / gate.FIXTURES / "valid/extra.csv").write_text(NS_HEADER + "\n", encoding="utf-8")
        self.assertIn(f"{gate.FIXTURES}/valid/extra.csv: no entry in {gate.EXPECTED}",
                      gate.run_check(self.root)["findings"])


if __name__ == "__main__":
    unittest.main()
