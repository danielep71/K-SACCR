#!/usr/bin/env python3
"""Business-day count on the EUR TARGET calendar, with a table of extra closing days (#721).

The Python reference of the day count of the workbook (Params: DaycountConvention,
YearConvention, Calendar). It holds three things:

- the TARGET closing days of a year, by rule, so that no yearly list has to be extended:
  1 January, Good Friday, Easter Monday, 1 May, 25 December and 26 December, with Good Friday
  and Easter Monday from the Gregorian Easter. Saturday and Sunday are not business days;
- the count of business days between two dates: the first date is excluded, the last is
  included, and an interval that runs backwards counts negative;
- the day count that turns a pair of dates into years: the ACTUAL days between them, counted
  on the workbook's calendar (``TARGET``, ``TARGET+TABLE`` or ``TABLE``), over the days of the
  conventional year (366, 365, 364, 360, 252 or 250). ``TABLE`` counts every day
  except those of the table, so calendar days while the table is empty.

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

# The days of the conventional year that YearConvention allows.
YEAR_BASES = (366.0, 365.0, 364.0, 360.0, 252.0, 250.0)
CALENDAR_TARGET = "TARGET"
CALENDAR_BOTH = "TARGET+TABLE"
CALENDAR_TABLE = "TABLE"
CALENDARS = (CALENDAR_TARGET, CALENDAR_BOTH, CALENDAR_TABLE)

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
        # The holidays come first, so that a date outside the range is refused on a weekend too.
        holidays = self.holidays_in(day.year)
        return day.weekday() < 5 and day not in holidays

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


def counted_days(start: dt.date, end: dt.date, calendar: str = CALENDAR_TABLE,
                 table: Iterable[dt.date] = ()) -> int:
    """The ACTUAL days in (start, end] that the workbook's calendar counts; negative when reversed.

    ``TARGET`` and ``TARGET+TABLE`` count business days (the weekend and the closing days are not
    counted). ``TABLE`` counts every day except those of the table, each once, weekends included.
    """
    if calendar not in CALENDARS:
        raise CalendarError(f"unknown calendar {calendar!r}: use one of {CALENDARS}")
    if end < start:
        return -counted_days(end, start, calendar, table)
    if calendar == CALENDAR_TARGET:
        return Calendar.target().business_days(start, end)
    if calendar == CALENDAR_BOTH:
        return Calendar.target(table).business_days(start, end)
    return (end - start).days - sum(1 for day in set(table) if start < day <= end)


def years_between(start: dt.date, end: dt.date, year_days: float = 365.0,
                  calendar: str = CALENDAR_TABLE, table: Iterable[dt.date] = ()) -> float:
    """The time from ``start`` to ``end`` in years: the counted days over the days of the year."""
    if year_days not in YEAR_BASES:
        raise CalendarError(f"unknown day-count convention {year_days!r}: use one of {YEAR_BASES}")
    return counted_days(start, end, calendar, table) / year_days
