"""Tests of the TARGET calendar and the business-day count (#721).

The calendar is checked against a second method (Gauss's Easter rule, written in a different form),
against dates fixed by the published TARGET closing days, and against a count made day by day.
"""

from __future__ import annotations

import datetime as dt
import random
import unittest

import business_calendar as bc
from business_calendar import Calendar, CalendarError


def easter_gauss(year: int) -> dt.date:
    """Easter by Gauss's rule with its two exceptions, valid for the years 1900 to 2099."""
    a, b, c = year % 19, year % 4, year % 7
    k = year // 100
    p = (13 + 8 * k) // 25
    q = k // 4
    m = (15 - p + k - q) % 30
    n = (4 + k - q) % 7
    d = (19 * a + m) % 30
    e = (2 * b + 4 * c + 6 * d + n) % 7
    day = 22 + d + e
    if day > 31:
        day = d + e - 9
        if day == 26:
            day = 19
        elif day == 25 and d == 28 and e == 6 and a > 10:
            day = 18
        return dt.date(year, 4, day)
    return dt.date(year, 3, day)


def count_by_day(calendar: Calendar, start: dt.date, end: dt.date) -> int:
    """Business days in (start, end] counted one day at a time."""
    count, day = 0, start
    while day < end:
        day += dt.timedelta(days=1)
        count += calendar.is_business_day(day)
    return count


class EasterTests(unittest.TestCase):
    def test_known_dates(self) -> None:
        known = {2000: (4, 23), 2008: (3, 23), 2011: (4, 24), 2024: (3, 31), 2025: (4, 20),
                 2026: (4, 5), 2027: (3, 28), 2028: (4, 16), 2038: (4, 25), 2100: (3, 28)}
        for year, (month, day) in known.items():
            self.assertEqual(bc.easter_sunday(year), dt.date(year, month, day), year)

    def test_second_method_agrees_from_1999_to_2099(self) -> None:
        for year in range(1999, 2100):
            self.assertEqual(bc.easter_sunday(year), easter_gauss(year), year)

    def test_always_a_sunday_between_22_march_and_25_april(self) -> None:
        for year in range(1999, 2600):
            easter = bc.easter_sunday(year)
            self.assertEqual(easter.weekday(), 6, year)
            self.assertTrue(dt.date(year, 3, 22) <= easter <= dt.date(year, 4, 25), year)

    def test_extremes_within_the_range(self) -> None:
        self.assertEqual(bc.easter_sunday(2008), dt.date(2008, 3, 23))
        self.assertEqual(bc.easter_sunday(2038), dt.date(2038, 4, 25))
        # The earliest date of the Gregorian cycle, 22 March, falls in 2285.
        self.assertEqual(bc.easter_sunday(2285), dt.date(2285, 3, 22))


class HolidayTests(unittest.TestCase):
    def test_published_closing_days(self) -> None:
        self.assertEqual(bc.target_holidays(2026), {
            dt.date(2026, 1, 1), dt.date(2026, 4, 3), dt.date(2026, 4, 6),
            dt.date(2026, 5, 1), dt.date(2026, 12, 25), dt.date(2026, 12, 26)})
        self.assertEqual(bc.target_holidays(2025), {
            dt.date(2025, 1, 1), dt.date(2025, 4, 18), dt.date(2025, 4, 21),
            dt.date(2025, 5, 1), dt.date(2025, 12, 25), dt.date(2025, 12, 26)})

    def test_the_years_before_the_full_rule(self) -> None:
        self.assertEqual(bc.target_holidays(1999),
                         {dt.date(1999, 1, 1), dt.date(1999, 12, 25), dt.date(1999, 12, 31)})
        self.assertEqual(bc.target_holidays(2000), {
            dt.date(2000, 1, 1), dt.date(2000, 4, 21), dt.date(2000, 4, 24),
            dt.date(2000, 5, 1), dt.date(2000, 12, 25), dt.date(2000, 12, 26)})
        self.assertIn(dt.date(2001, 12, 31), bc.target_holidays(2001))
        self.assertNotIn(dt.date(2002, 12, 31), bc.target_holidays(2002))

    def test_every_year_has_six_closing_days_from_2000(self) -> None:
        for year in range(2002, 2300):
            holidays = bc.target_holidays(year)
            self.assertEqual(len(holidays), 6, year)
            easter = bc.easter_sunday(year)
            self.assertIn(easter - dt.timedelta(days=2), holidays)
            self.assertIn(easter + dt.timedelta(days=1), holidays)

    def test_before_the_first_year_is_refused(self) -> None:
        with self.assertRaises(CalendarError):
            bc.target_holidays(1998)
        with self.assertRaises(CalendarError):
            Calendar.target().business_days(dt.date(1998, 12, 1), dt.date(1999, 2, 1))


class BusinessDayTests(unittest.TestCase):
    def setUp(self) -> None:
        self.target = Calendar.target()

    def test_same_day_and_adjacent_days(self) -> None:
        day = dt.date(2026, 6, 30)  # a Tuesday
        self.assertEqual(self.target.business_days(day, day), 0)
        self.assertEqual(self.target.business_days(day, dt.date(2026, 7, 1)), 1)

    def test_across_a_weekend(self) -> None:
        self.assertEqual(self.target.business_days(dt.date(2026, 7, 3), dt.date(2026, 7, 6)), 1)
        self.assertEqual(self.target.business_days(dt.date(2026, 7, 4), dt.date(2026, 7, 5)), 0)

    def test_across_easter_and_christmas(self) -> None:
        # Thursday 2 April to Tuesday 7 April 2026: Good Friday and Easter Monday are closed.
        self.assertEqual(self.target.business_days(dt.date(2026, 4, 2), dt.date(2026, 4, 7)), 1)
        # Wednesday 24 December to Monday 28 December 2026: 25th is Friday, 26th a Saturday.
        self.assertEqual(self.target.business_days(dt.date(2026, 12, 24), dt.date(2026, 12, 28)), 1)

    def test_a_year(self) -> None:
        # Worked by hand: weekdays of the year less the closing days that fall on a weekday.
        expected = {2024: 256, 2025: 255, 2026: 256, 2027: 258, 2028: 255}
        for year, business_days in expected.items():
            self.assertEqual(self.target.business_days(dt.date(year - 1, 12, 31),
                                                       dt.date(year, 12, 31)), business_days, year)

    def test_every_year_against_the_second_method(self) -> None:
        counts = set()
        for year in range(2000, 2100):
            easter = easter_gauss(year)
            closed = {dt.date(year, 1, 1), easter - dt.timedelta(days=2), easter + dt.timedelta(days=1),
                      dt.date(year, 5, 1), dt.date(year, 12, 25), dt.date(year, 12, 26)}
            if year == 2001:
                closed.add(dt.date(2001, 12, 31))
            day, expected = dt.date(year, 1, 1), 0
            while day.year == year:
                expected += day.weekday() < 5 and day not in closed
                day += dt.timedelta(days=1)
            got = self.target.business_days(dt.date(year - 1, 12, 31), dt.date(year, 12, 31))
            self.assertEqual(got, expected, year)
            counts.add(got)
        self.assertEqual((min(counts), max(counts)), (254, 259))

    def test_reversed_interval_is_negative(self) -> None:
        start, end = dt.date(2026, 6, 30), dt.date(2027, 6, 30)
        self.assertEqual(self.target.business_days(end, start), -self.target.business_days(start, end))

    def test_matches_a_count_day_by_day(self) -> None:
        generator = random.Random(721)
        origin = dt.date(1999, 1, 1)
        extra = Calendar.target({dt.date(2026, 8, 15), dt.date(2026, 12, 8), dt.date(2027, 1, 6)})
        for calendar in (self.target, extra, Calendar.from_table({dt.date(2026, 8, 12)})):
            for _ in range(300):
                start = origin + dt.timedelta(days=generator.randrange(0, 12000))
                end = start + dt.timedelta(days=generator.randrange(0, 1500))
                self.assertEqual(calendar.business_days(start, end),
                                 count_by_day(calendar, start, end), (start, end))


class TableTests(unittest.TestCase):
    def test_extra_day_is_counted_once_and_a_weekend_day_not_at_all(self) -> None:
        extra = Calendar.target({dt.date(2026, 8, 14), dt.date(2026, 8, 15), dt.date(2026, 8, 14)})
        # 14 August 2026 is a Friday, 15 August a Saturday.
        self.assertEqual(extra.business_days(dt.date(2026, 8, 13), dt.date(2026, 8, 17)), 1)
        self.assertEqual(Calendar.target().business_days(dt.date(2026, 8, 13), dt.date(2026, 8, 17)), 2)

    def test_the_table_alone_has_no_target_days(self) -> None:
        table = Calendar.from_table({dt.date(2026, 12, 8)})
        self.assertTrue(table.is_business_day(dt.date(2026, 12, 25)))
        self.assertFalse(table.is_business_day(dt.date(2026, 12, 8)))
        self.assertEqual(table.business_days(dt.date(2026, 12, 7), dt.date(2026, 12, 9)), 1)

    def test_the_table_has_no_range_of_validity(self) -> None:
        table = Calendar.from_table(())
        self.assertEqual(table.business_days(dt.date(1990, 1, 1), dt.date(1990, 1, 8)), 5)

    def test_fingerprint_follows_the_rule_and_the_table(self) -> None:
        plain = Calendar.target()
        self.assertEqual(plain.fingerprint(), Calendar.target(()).fingerprint())
        self.assertEqual(len(plain.fingerprint()), 8)
        one = Calendar.target({dt.date(2026, 8, 15)})
        two = Calendar.target({dt.date(2026, 8, 15), dt.date(2026, 12, 8)})
        self.assertEqual(len({plain.fingerprint(), one.fingerprint(), two.fingerprint(),
                              Calendar.from_table(()).fingerprint()}), 4)
        self.assertEqual(one.fingerprint(), Calendar.target([dt.date(2026, 8, 15)]).fingerprint())


class ConventionTests(unittest.TestCase):
    def test_calendar_days_over_the_days_per_year(self) -> None:
        start, end = dt.date(2026, 6, 30), dt.date(2028, 5, 15)
        self.assertEqual(bc.years_between(start, end), (end - start).days / 365)
        self.assertEqual(bc.years_between(start, end, days_per_year=362), (end - start).days / 362)

    def test_business_days_over_the_business_days_per_year(self) -> None:
        start, end = dt.date(2026, 6, 30), dt.date(2028, 5, 15)
        expected = Calendar.target().business_days(start, end) / 250
        self.assertEqual(bc.years_between(start, end, bc.BD250), expected)
        self.assertEqual(bc.years_between(start, end, bc.BD250, business_days_per_year=252),
                         Calendar.target().business_days(start, end) / 252)

    def test_a_business_day_year_is_longer_than_a_calendar_year_on_long_maturities(self) -> None:
        start = dt.date(2026, 6, 30)
        end = dt.date(2036, 6, 30)
        self.assertGreater(bc.years_between(start, end, bc.BD250), bc.years_between(start, end))

    def test_unknown_convention_is_refused(self) -> None:
        with self.assertRaises(CalendarError):
            bc.years_between(dt.date(2026, 1, 1), dt.date(2026, 2, 1), "ACT360")


if __name__ == "__main__":
    unittest.main()
