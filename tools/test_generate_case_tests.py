"""Tests for the generator of tests/modules/TEST_Cases.bas."""
import json
import shutil
import tempfile
import unittest
from pathlib import Path

import generate_case_tests as generator

ROOT = Path(__file__).resolve().parents[1]


class GenerateCaseTestsTests(unittest.TestCase):
    def test_committed_module_is_current(self) -> None:
        committed = (ROOT / generator.MODULE).read_bytes().decode("cp1252")
        self.assertEqual(committed, generator.generate(ROOT))

    def test_numbers_are_plain_decimal_text(self) -> None:
        self.assertEqual(generator.text(1e-09), '"0.000000001"')
        self.assertEqual(generator.text(5845000000), '"5845000000"')
        self.assertEqual(generator.text(-0.27), '"-0.27"')
        self.assertEqual(generator.text(2.0), '"2"')
        self.assertEqual(generator.text(None), '""')
        self.assertEqual(generator.text(True), '"Y"')
        self.assertEqual(generator.text('say "hi"'), '"say ""hi"""')

    def test_non_ascii_text_is_rejected(self) -> None:
        with self.assertRaises(ValueError):
            generator.text("café")

    def test_every_expected_file_becomes_one_case(self) -> None:
        module = generator.generate(ROOT)
        expected_files = sorted((ROOT / generator.EXPECTED).glob("*.json"))
        outputs = sum(len(json.loads(p.read_text())["outputs"]) for p in expected_files)
        self.assertIn(f'If Not TEST_CaseRunner.BeginSuite("TEST_Cases", {len(expected_files)}, {outputs}) '
                      'Then Exit Sub\r\n', module)
        self.assertEqual(module.count("TEST_CaseRunner.Expect"), outputs)
        self.assertNotIn("\n", module.replace("\r\n", ""))

    def test_fixture_parameters_are_set_after_begin_case(self) -> None:
        module = generator.generate(ROOT).split("\r\n")
        begin = module.index('        TEST_CaseRunner.BeginCase "cre99-example-3", "BCBS", "2026-09-30", "USD"')
        self.assertEqual(module[begin + 1], '        TEST_CaseRunner.SetParameter "DaysPerYear", "360"')

    def test_check_mode_detects_a_stale_module(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            for folder in (generator.FIXTURES, generator.EXPECTED):
                shutil.copytree(ROOT / folder, root / folder)
            (root / "tests/modules").mkdir(parents=True)
            (root / generator.MODULE).write_bytes(generator.generate(ROOT).encode("cp1252"))
            self.assertEqual(generator.generate(root), generator.generate(ROOT))
            case = root / generator.EXPECTED / "cre99-example-1.bcbs.json"
            data = json.loads(case.read_text())
            data["outputs"][0]["value"] = 1
            case.write_text(json.dumps(data))
            self.assertNotEqual(generator.generate(root),
                                (root / generator.MODULE).read_bytes().decode("cp1252"))


if __name__ == "__main__":
    unittest.main()
