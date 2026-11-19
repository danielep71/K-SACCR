"""The measured differences of an Excel run against the accuracy contracts (#537); no test runs Excel."""
from __future__ import annotations

import contextlib
import io
import tempfile
import unittest
from pathlib import Path

import measured_differences as measured

ROOT = Path(__file__).resolve().parents[1]


def line(case: str, quantity: str, expected: float, difference: float, trade: str = "") -> str:
    return (f"MEASURED={case}; trade={trade}; quantity={quantity}; expected={expected!r}; "
            f"actual={expected + difference!r}; difference={difference!r}")


class MeasuredDifferencesTests(unittest.TestCase):
    def setUp(self) -> None:
        self.document, self.expected = measured.load(ROOT)
        required, _ = measured.outputs(self.document, self.expected)
        self.required = required

    def perfect(self) -> list[str]:
        return [line(key.case, key.quantity, wanted.value, 0.0, key.trade)
                for key, wanted in sorted(self.required.items())]

    def problems(self, lines: list[str]) -> list[str]:
        return measured.check(self.document, self.expected, "\n".join(lines))[0]

    def test_a_perfect_run_has_no_problem_and_every_family(self) -> None:
        problems, families = measured.check(self.document, self.expected, "\n".join(self.perfect()))
        self.assertEqual(problems, [])
        self.assertEqual(set(families), {"EAD", "RC", "PFE", "add-ons"})
        self.assertEqual(sum(row["outputs"] for row in families.values()), len(self.required))

    def test_full_precision_cases_allow_the_claim_and_no_more(self) -> None:
        key = next(k for k in self.required if k.case == "illustrative-ccp-non-qualifying.crr"
                   and k.quantity == "exposure_value")
        value = self.required[key].value
        lines = self.perfect()
        position = next(i for i, text in enumerate(lines) if f"; quantity=exposure_value; expected={value!r}" in text
                        and text.startswith(f"MEASURED={key.case};"))
        permitted = self.required[key].permitted
        self.assertLess(permitted, 1e-3)
        lines[position] = line(key.case, key.quantity, value, permitted * 0.5)
        self.assertEqual(self.problems(lines), [])
        lines[position] = line(key.case, key.quantity, value, permitted * 3)
        self.assertEqual(len(self.problems(lines)), 1)
        self.assertIn("exceeds the contract", self.problems(lines)[0])

    def test_a_rounded_case_is_held_to_its_rounding(self) -> None:
        case = "cre99-example-1.bcbs"
        key = next(k for k in self.required if k.case == case and k.quantity == "exposure_value")
        lines = self.perfect()
        position = next(i for i, text in enumerate(lines) if text.startswith(f"MEASURED={case}; trade=; quantity=exposure_value;"))
        value = self.required[key].value
        self.assertEqual(self.required[key].permitted, 500.0)
        lines[position] = line(case, key.quantity, value, 499.0)
        self.assertEqual(self.problems(lines), [])
        lines[position] = line(case, key.quantity, value, 501.0)
        self.assertIn("exceeds the contract", " ".join(self.problems(lines)))

    def test_missing_extra_and_unreadable_lines(self) -> None:
        lines = self.perfect()
        self.assertIn("no measured difference", " ".join(self.problems(lines[1:])))
        self.assertIn("measured twice", " ".join(self.problems(lines + lines[:1])))
        self.assertIn("unreadable", " ".join(self.problems(lines + ["MEASURED=garbage"])))
        self.assertIn("not a finite number", " ".join(self.problems(lines + ["MEASURED=a.b; trade=; quantity=q; "
                                                                           "expected=1; actual=nan; difference=0"])))

    def test_lines_of_other_quantities_are_ignored(self) -> None:
        extra = line("cre99-example-1.bcbs", "adjusted_notional", 78694000.0, 3.0, "EX1-T1")
        self.assertEqual(self.problems(self.perfect() + [extra, "RESULT=PASS"]), [])

    def test_cli_prints_the_table_and_the_exit_code(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            log = Path(folder) / "TEST_Cases.log"
            log.write_text("\n".join(self.perfect()), encoding="utf-8")
            with contextlib.redirect_stdout(io.StringIO()) as out:
                self.assertEqual(measured.main(["--root", str(ROOT), "--log", str(log)]), 0)
            self.assertIn("| EAD |", out.getvalue())
            log.write_text("\n".join(self.perfect()[1:]), encoding="utf-8")
            with contextlib.redirect_stdout(io.StringIO()) as out:
                self.assertEqual(measured.main(["--root", str(ROOT), "--log", str(log)]), 1)
            self.assertIn("no measured difference", out.getvalue())
            with contextlib.redirect_stderr(io.StringIO()):
                self.assertEqual(measured.main(["--root", str(ROOT), "--log", str(Path(folder) / "none.log")]), 2)


if __name__ == "__main__":
    unittest.main()
