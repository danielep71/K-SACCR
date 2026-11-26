"""The reference of the formatted run report (#259): its text rules and its row rules."""
from __future__ import annotations

import unittest

import report


class TextTests(unittest.TestCase):
    def test_the_examples_of_the_specification(self) -> None:
        for example in report.load_spec()["formats"]["examples"]:
            with self.subTest(example=example):
                self.assertEqual(report.cell_text(example["kind"], example["value"]), example["text"])

    def test_an_amount_rounds_half_away_from_zero_on_its_15_digit_form(self) -> None:
        # 2.675 is 2.67499999999999982236431605997495353221893310546875 as a double: its 15-digit form rounds up.
        self.assertEqual(report.number_text(2.675, 2, True), "2.68")
        self.assertEqual(report.number_text(-2.675, 2, True), "-2.68")
        self.assertEqual(report.number_text(1.005, 2, True), "1.01")
        self.assertEqual(report.number_text(0.125, 2, True), "0.13")
        self.assertEqual(report.number_text(-0.125, 2, True), "-0.13")

    def test_a_value_that_rounds_to_zero_has_no_sign(self) -> None:
        self.assertEqual(report.number_text(-0.004, 2, True), "0.00")
        self.assertEqual(report.number_text(-0.0000004, 6, False), "0.000000")
        self.assertEqual(report.number_text(-0.0000005, 6, False), "-0.000001")

    def test_thousands_are_grouped_by_three_with_a_comma(self) -> None:
        self.assertEqual(report.number_text(999999.995, 2, True), "1,000,000.00")
        self.assertEqual(report.number_text(-1234567890123.456, 2, True), "-1,234,567,890,123.46")
        self.assertEqual(report.number_text(100, 2, True), "100.00")
        self.assertEqual(report.number_text(1000, 0, True), "1,000")
        self.assertEqual(report.number_text(1234567.5, 0, False), "1234568")

    def test_a_blank_stays_blank_and_a_text_stays_a_text(self) -> None:
        self.assertEqual(report.cell_text("amount", None), "")
        self.assertEqual(report.cell_text("amount", ""), "")
        self.assertEqual(report.cell_text("text", "n/a"), "n/a")

    def test_dates(self) -> None:
        self.assertEqual(report.date_text(61), "1900-03-01")
        self.assertEqual(report.date_text(46295), "2026-09-30")
        self.assertEqual(report.date_text(46295.9999), "2026-09-30")  # a time of day is ignored
        self.assertEqual(report.date_text(73415), "2100-12-31")
        with self.assertRaises(report.ReportError):
            report.date_text(60)

    def test_a_time_rounds_to_the_second_and_may_carry_into_the_next_day(self) -> None:
        self.assertEqual(report.datetime_text(46352 + 0.6 / 86400), "2026-11-26 00:00:01")
        self.assertEqual(report.datetime_text(46352 + 86399.6 / 86400), "2026-11-27 00:00:00")
        self.assertEqual(report.datetime_text(46352.75), "2026-11-26 18:00:00")

    def test_an_unknown_kind_is_refused(self) -> None:
        with self.assertRaises(report.ReportError):
            report.cell_text("percent", 1)


CHECKS = [
    ("ERROR", "Trades", 39, "NIR-B", "price", "RUN", "TR-SHIFTED-PRICE", ""),
    ("WARNING", "Trades", 12, "A", "late", "RUN", "TR-LATE", ""),
    ("ERROR", "Trades", 44, "OT-X", "other", "RUN", "TR-OTHER-REGIME", "Asset class"),
    ("INFO", "Checks", None, "", "note", "RUN", "INF-NOTE", ""),
    ("ERROR", "Trades", 56, "A11", "matured", "RUN", "TR-MATURED", ""),
    ("WARNING", "Trades", 13, "B", "late", "RUN", "TR-LATE", ""),
    ("ERROR", "Trades", 40, "NIR-C", "price", "RUN", "TR-SHIFTED-PRICE", ""),
    ("WARNING", "NettingSets", 5, "N", "no code", "RUN", "", ""),
    ("ERROR", "Trades", 60, "Z", "late but error", "RUN", "TR-LATE", ""),
]


class ChecksTests(unittest.TestCase):
    def test_the_summary_counts_the_rows_of_each_code_with_its_highest_severity(self) -> None:
        self.assertEqual(report.checks_by_code(CHECKS), [
            ("TR-LATE", "ERROR", 3),
            ("TR-MATURED", "ERROR", 1),
            ("TR-OTHER-REGIME", "ERROR", 1),
            ("TR-SHIFTED-PRICE", "ERROR", 2),
            ("(none)", "WARNING", 1),
            ("INF-NOTE", "INFO", 1),
        ])

    def test_the_rows_follow_the_summary_and_keep_their_order_within_a_code(self) -> None:
        ordered = report.checks_in_code_order(CHECKS)
        self.assertEqual([row[6] for row in ordered], [
            "TR-LATE", "TR-LATE", "TR-LATE", "TR-MATURED", "TR-OTHER-REGIME",
            "TR-SHIFTED-PRICE", "TR-SHIFTED-PRICE", "", "INF-NOTE"])
        self.assertEqual([row[2] for row in ordered if row[6] == "TR-LATE"], [12, 13, 60])
        self.assertEqual(len(ordered), len(CHECKS))

    def test_no_row_means_no_code(self) -> None:
        self.assertEqual(report.checks_by_code([]), [])
        self.assertEqual(report.checks_in_code_order([]), [])

    def test_a_clean_run_has_one_info_code(self) -> None:
        clean = [("INFO", "Checks", None, "", "No findings", "RUN", "RUN-CLEAN", "")]
        self.assertEqual(report.checks_by_code(clean), [("RUN-CLEAN", "INFO", 1)])

    def test_a_blank_code_and_a_code_of_spaces_are_the_same_code(self) -> None:
        rows = [("ERROR", "S", 1, "", "m", "RUN", "  ", ""), ("ERROR", "S", 2, "", "m", "RUN", None, "")]
        self.assertEqual(report.checks_by_code(rows), [("(none)", "ERROR", 2)])


class ReconciliationTests(unittest.TestCase):
    ROWS = [("A", "MATCH"), ("B", "DIFFERENCE"), ("C", "match"), ("D", "MISSING IN CALCULATOR"),
            ("E", "NOT IN BANK FILE"), ("F", "MATCH")]

    def test_the_pdf_keeps_every_row_that_is_not_a_match(self) -> None:
        kept, omitted = report.pdf_reconciliation_rows(self.ROWS, 1)
        self.assertEqual([row[0] for row in kept], ["B", "D", "E"])
        self.assertEqual(omitted, 3)

    def test_nothing_is_left_out_when_no_row_matches(self) -> None:
        kept, omitted = report.pdf_reconciliation_rows([("A", "DIFFERENCE")], 1)
        self.assertEqual((len(kept), omitted), (1, 0))
        self.assertEqual(report.omitted_note(0), "")

    def test_the_note_counts_the_rows_left_out(self) -> None:
        self.assertEqual(report.omitted_note(1),
                         "1 row with the result MATCH is left out; the workbook version of the report has every row.")
        self.assertEqual(report.omitted_note(3),
                         "3 rows with the result MATCH are left out; the workbook version of the report has every row.")


class SpecificationTests(unittest.TestCase):
    def test_the_specification_is_sound(self) -> None:
        self.assertEqual(report.check_spec(report.load_spec()), [])

    def test_a_broken_specification_is_reported(self) -> None:
        spec = report.load_spec()
        spec["sections"][3]["columns"][0]["kind"] = "percent"
        spec["sections"][4]["sheet"] = "Results"
        spec["formats"]["examples"][0]["text"] = "1234567.89"
        problems = report.check_spec(spec)
        self.assertEqual(len(problems), 3, problems)


if __name__ == "__main__":
    unittest.main()
