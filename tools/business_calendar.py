#!/usr/bin/env python3
"""Business-day count on the EUR TARGET calendar, with a table of extra closing days (#721).

The Python reference of the day-count convention that the workbook will offer next to the
calendar-day year. It holds three things:

- the TARGET closing days of a year, by rule, so that no yearly list has to be extended:
  1 January, Good Friday, Easter Monday, 1 May, 25 December and 26 December, with Good Friday
  and Easter Monday from the Gregorian Easter. Saturday and Sunday are not business days;
- the count of business days between two dates: the first date is excluded, the last is
  included, and an interval that runs backwards counts negative;
- the two conventions that turn a pair of dates into years: ``CAL365`` (calendar days over
  ``DaysPerYear``, the behaviour of today) and ``BD250`` (business days over
  ``BusinessDaysPerYear``).

A :class:`Calendar` is either the TARGET rule, optionally with extra closing days, or a table
of closing days alone. Weekends are never counted twice, so a table may list a holiday that
falls on one. :meth:`Calendar.fingerprint` identifies the calendar for the evidence of a run.

Range of validity of the TARGET rule: the system opened on 1 January 1999. In 1999 its closing
days were 1 January, 25 December and 31 December. Good Friday, Easter Monday, 1 May and
26 December closed it from 2000, and 31 December once more in 2001. Dates before 1 January 1999
are refused. There is no upper limit: the Gregorian Easter rule is exact for every year of the
Gregorian calendar.
"""

from __future__ import annotations

import datetime as dt
import hashlib
from collections.abc import Iterable
from dataclasses import dataclass

CAL365 = "CAL365"
BD250 = "BD250"
CONVENTIONS = (CAL365, BD250)

# The first year of the TARGET calendar.
FIRST_YEAR = 1999
# From this year the system also closes on Good Friday, Easter Monday, 1 May and 26 December.
FULL_RULE_YEAR = 2000
# The years in which it also closed on 31 December.
NEW_YEAR_EVE_YEARS = frozenset({1998, 1999, 2001})


class CalendarError(ValueError):
    """A date outside the range of validity of a calendar, or an unknown convention."""


def easter_sunday(year: int) -> dt.date:
    """The date of Easter Sunday in the Gregorian calendar (the Anonymous Gregorian algorithm)."""
    a = year % 19
    b, c = divmod(year, 100)
    d, e = divmod(b, 4)
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i, k = divmod(c, 4)
    el = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * el) // 451
    month, day = divmod(h + el - 7 * m + 114, 31)
    return dt.date(year, month, day + 1)


def target_holidays(year: int) -> frozenset[dt.date]:
    """The days on which TARGET is closed in a year, weekends included when they are holidays."""
    if year < FIRST_YEAR:
        raise CalendarError(f"TARGET opened in {FIRST_YEAR}: year {year} is outside its range")
    days = {dt.date(year, 1, 1), dt.date(year, 12, 25)}
    if year >= FULL_RULE_YEAR:
        easter = easter_sunday(year)
        days |= {easter - dt.timedelta(days=2), easter + dt.timedelta(days=1),
                 dt.date(year, 5, 1), dt.date(year, 12, 26)}
    if year in NEW_YEAR_EVE_YEARS:
        days.add(dt.date(year, 12, 31))
    return frozenset(days)


def weekdays_through(day: dt.date) -> int:
    """The number of Mondays to Fridays from 1 January of year 1 up to and including ``day``."""
    ordinal = day.toordinal()  # ordinal 1 is a Monday
    return 5 * (ordinal // 7) + min(ordinal % 7, 5)


@dataclass(frozen=True)
class Calendar:
    """The closing days of a calendar: the TARGET rule and/or a table of extra days."""

    use_target: bool
    table: frozenset[dt.date]

    @staticmethod
    def target(extra: Iterable[dt.date] = ()) -> Calendar:
        """The TARGET rule, with extra closing days if any."""
        return Calendar(True, frozenset(extra))

    @staticmethod
    def from_table(days: Iterable[dt.date]) -> Calendar:
        """A calendar made of the weekend and of the days of the table only."""
        return Calendar(False, frozenset(days))

    def holidays_in(self, year: int) -> frozenset[dt.date]:
        """The closing days of a year, from the rule and from the table."""
        rule = target_holidays(year) if self.use_target else frozenset()
        return rule | {day for day in self.table if day.year == year}

    def is_business_day(self, day: dt.date) -> bool:
        return day.weekday() < 5 and day not in self.holidays_in(day.year)

    def business_days(self, start: dt.date, end: dt.date) -> int:
        """Business days in (start, end]; negative when ``end`` is before ``start``."""
        if end < start:
            return -self.business_days(end, start)
        count = weekdays_through(end) - weekdays_through(start)
        for year in range(start.year, end.year + 1):
            count -= sum(1 for day in self.holidays_in(year)
                         if day.weekday() < 5 and start < day <= end)
        return count

    def fingerprint(self) -> str:
        """Eight hexadecimal characters that identify the rule and the table."""
        text = ("TARGET" if self.use_target else "TABLE") + ";" \
            + ";".join(day.isoformat() for day in sorted(self.table))
        return hashlib.sha256(text.encode("ascii")).hexdigest()[:8].upper()


def years_between(start: dt.date, end: dt.date, convention: str = CAL365,
                  calendar: Calendar | None = None, days_per_year: float = 365,
                  business_days_per_year: float = 250) -> float:
    """The time from ``start`` to ``end`` in years on one of the two conventions."""
    if convention == CAL365:
        return (end - start).days / days_per_year
    if convention == BD250:
        return (calendar or Calendar.target()).business_days(start, end) / business_days_per_year
    raise CalendarError(f"unknown day-count convention {convention!r}: use one of {CONVENTIONS}")
