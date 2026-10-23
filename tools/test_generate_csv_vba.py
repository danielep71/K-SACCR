"""Tests of tools/generate_csv_vba.py, the generator of the VBA side of the CSV contract (#207)."""
import copy
import unittest
from collections.abc import Callable
from pathlib import Path
from typing import Any

import generate_csv_vba as gen
from csv_contract import load_contract

ROOT = Path(__file__).resolve().parents[1]


class GeneratorTests(unittest.TestCase):
    def test_committed_modules_are_current(self) -> None:
        for path, content in gen.generate(ROOT).items():
            with self.subTest(path=path):
                committed = (ROOT / path).read_bytes().decode("cp1252").replace("\r\n", "\n")
                self.assertEqual(committed, content.replace("\r\n", "\n"))

    def test_modules_are_crlf_and_fit_a_vba_line(self) -> None:
        for path, content in gen.generate(ROOT).items():
            with self.subTest(path=path):
                self.assertNotIn("\n", content.replace("\r\n", ""))
                self.assertLess(max(len(line) for line in content.split("\r\n")), 1024)

    def test_strings_outside_printable_ascii_become_chrw(self) -> None:
        self.assertEqual(gen.vba_string('a"b'), '"a""b"')
        self.assertEqual(gen.vba_string("é"), "ChrW(233)")
        self.assertEqual(gen.vba_string("x\U0001F600"), '"x" & ChrW(55357) & ChrW(56832)')
        self.assertEqual(gen.vba_string(""), '""')

    def test_an_unsupported_contract_is_refused(self) -> None:
        contract = load_contract(ROOT)
        changes: tuple[Callable[[dict[str, Any]], None], ...] = (
            lambda c: c["types"].update(uuid={}),
            lambda c: c["types"]["currency"].update(pattern="^[A-Z]{4}$"),
            lambda c: c["files"]["Trades"]["row_rules"].append({"id": "x", "below": ["a", "b"]}))
        for change in changes:
            changed = copy.deepcopy(contract)
            change(changed)
            with self.assertRaises(ValueError):
                gen.schema_module(changed)


if __name__ == "__main__":
    unittest.main()
