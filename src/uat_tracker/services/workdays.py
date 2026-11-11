"""Working days (PART2 §B): Mon-Fri minus cached bank holidays."""

from __future__ import annotations

import datetime as dt


class NoHolidayData(Exception):
    """The holiday cache has nothing for a year we need, so workdays can't be trusted."""


class WorkCalendar:
    def __init__(self, holidays: set[dt.date], known_years: set[int]):
        self.holidays = set(holidays)
        self.known_years = set(known_years)

    @classmethod
    def from_repo(cls, repo) -> "WorkCalendar":
        return cls({dt.date.fromisoformat(d) for d in repo.holidays()}, repo.holiday_years())

    def has_year(self, year: int) -> bool:
        return year in self.known_years

    def is_workday(self, d: dt.date) -> bool:
        return d.weekday() < 5 and d not in self.holidays

    def on_or_before(self, d: dt.date) -> dt.date:
        """Trigger date on a non-workday -> previous workday."""
        while not self.is_workday(d):
            d -= dt.timedelta(days=1)
        return d

    def workdays_between(self, start: dt.date, end: dt.date) -> list[dt.date]:
        """Workdays with start < d < end."""
        out, d = [], start + dt.timedelta(days=1)
        while d < end:
            if self.is_workday(d):
                out.append(d)
            d += dt.timedelta(days=1)
        return out
