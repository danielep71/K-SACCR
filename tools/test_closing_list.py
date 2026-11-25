"""The engine's list of closing days (#771) against the calendar's reference.

CORE_Engine counts the business days of a date on a business-day convention as
the weekdays between the reporting date and the date less the closing weekdays
between them, both from one sorted list built once: the TARGET closing days of
1999 to 2100 and the days of the table. This model builds the list as
LoadClosingDays does, in VBA serial numbers, and compares the count with
business_calendar.counted_days, the reference of SACCR_BusinessDays (#721), on
every day of the window and on random intervals, with and without a table.
TEST_Formulas makes the same comparison in Excel against SACCR_BusinessDays.
"""

from __future__ import annotations

import bisect
import datetime as dt
import random
import unittest
from pathlib import Path

import business_calendar as bc

ROOT = Path(__file__).resolve().parents[1]
EPOCH = dt.date(1899, 12, 30)  # serial 0 of VBA's Date
WINDOW_FIRST = 36161  # 1 January 1999
WINDOW_LAST = 73415  # 31 December 2100


def serial(day: dt.date) -> int:
    return (day - EPOCH).days


def date_of(number: int) -> dt.date:
    return EPOCH + dt.timedelta(days=number)


def is_weekday(number: int) -> bool:
    return (number - 2) % 7 <= 4


def weekdays_through(number: int) -> int:
    offset = number - 2
    return 5 * (offset // 7) + min(offset % 7 + 1, 5)


def target_days() -> list[int]:
    """LoadClosingDays' TARGET days, in the order it writes them."""
    days = []
    for year in range(1999, 2101):
        days.append(serial(dt.date(year, 1, 1)))
        if year >= 2000:
            easter = serial(bc.easter_sunday(year))
            days += [easter - 2, easter + 1, serial(dt.date(year, 5, 1))]
        days.append(serial(dt.date(year, 12, 25)))
        if year >= 2000:
            days.append(serial(dt.date(year, 12, 26)))
        if year in (1999, 2001):
            days.append(serial(dt.date(year, 12, 31)))
    return days


def closing_list(table: list[float]) -> list[int]:
    """The list LoadClosingDays builds: TARGET and the table's days in the window, weekdays, once."""
    table_days = sorted(int(item) for item in table if WINDOW_FIRST <= item < WINDOW_LAST + 1)
    merged: list[int] = []
    for day in sorted(target_days() + table_days):
        if is_weekday(day) and (not merged or merged[-1] != day):
            merged.append(day)
    return merged


def list_count(closing: list[int], first: int, last: int) -> int:
    return (weekdays_through(last) - weekdays_through(first)) \
        - (bisect.bisect_right(closing, last) - bisect.bisect_right(closing, first))


TABLE = [serial(dt.date(2030, 6, 3)), serial(dt.date(2030, 6, 1)),  # a Monday; a Saturday
         serial(dt.date(2026, 12, 25)), serial(dt.date(2026, 12, 25)),  # a TARGET day, twice
         serial(dt.date(2024, 8, 15)) + 0.5,  # a fraction of a day
         serial(dt.date(1998, 6, 1)), serial(dt.date(2101, 3, 1)),  # outside the window
         serial(dt.date(2099, 11, 2)), serial(dt.date(2000, 1, 3))]  # unsorted


class ClosingListTests(unittest.TestCase):
    def test_target_days_are_written_in_date_order(self) -> None:
        days = target_days()
        self.assertEqual(days, sorted(days))
        self.assertLessEqual(len(days), 7 * 102)

    def check_every_day(self, calendar: str, table: list[float]) -> None:
        closing = closing_list(table)
        reference_table = [date_of(int(item)) for item in table]
        for last in range(WINDOW_FIRST + 1, WINDOW_LAST + 1):
            expected = bc.counted_days(date_of(last - 1), date_of(last), calendar, reference_table)
            self.assertEqual(list_count(closing, last - 1, last), expected, date_of(last))

    def test_every_day_on_target(self) -> None:
        self.check_every_day(bc.CALENDAR_TARGET, [])

    def test_every_day_on_target_and_a_table(self) -> None:
        self.check_every_day(bc.CALENDAR_BOTH, TABLE)

    def test_random_intervals_and_the_edges_of_the_window(self) -> None:
        rng = random.Random(771)
        pairs = [(WINDOW_FIRST, WINDOW_LAST), (WINDOW_FIRST, WINDOW_FIRST), (WINDOW_LAST - 1, WINDOW_LAST)]
        for _ in range(5000):
            first = rng.randrange(WINDOW_FIRST, WINDOW_LAST + 1)
            pairs.append((first, rng.randrange(first, min(first + 40 * 366, WINDOW_LAST) + 1)))
        for calendar, table in ((bc.CALENDAR_TARGET, []), (bc.CALENDAR_BOTH, TABLE)):
            closing = closing_list(table)
            reference_table = [date_of(int(item)) for item in table]
            for first, last in pairs:
                expected = bc.counted_days(date_of(first), date_of(last), calendar, reference_table)
                self.assertEqual(list_count(closing, first, last), expected, (calendar, first, last))

    def test_the_engine_holds_this_rule(self) -> None:
        source = (ROOT / "src" / "core" / "CORE_Engine.bas").read_text(encoding="cp1252")
        body = source[source.index("Private Sub LoadClosingDays("):]
        body = body[:body.index("End Sub")]
        for line in ("Private Const BD_WINDOW_FIRST        As Long = 36161",
                     "Private Const BD_WINDOW_LAST         As Long = 73415",
                     "Private Const BD_TARGET_FIRST_YEAR   As Long = 1999",
                     "Private Const BD_TARGET_FULL_YEAR    As Long = 2000"):
            self.assertIn(line, source)
        for line in ("TargetDays(TargetCount + 1) = Easter - 2",
                     "TargetDays(TargetCount + 2) = Easter + 1",
                     "TargetDays(TargetCount + 3) = CLng(DateSerial(Yr, 5, 1))",
                     "TargetDays(TargetCount) = CLng(DateSerial(Yr, 12, 26))",
                     "If Yr = 1999 Or Yr = 2001 Then",
                     "If ((Closing - 2) Mod 7) <= 4 Then"):
            self.assertIn(line, body)


if __name__ == "__main__":
    unittest.main()
