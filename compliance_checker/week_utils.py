"""
Week calculation helpers. Weeks run Mon–Sun, Europe/London timezone.
"""
from __future__ import annotations

from datetime import datetime, timedelta, date
import pytz

LONDON = pytz.timezone("Europe/London")


def week_bounds(ref: date | datetime | str = None) -> tuple[datetime, datetime]:
    """
    Return (week_start, week_end) as timezone-aware datetimes (Europe/London).

    ref can be:
      - None → current week
      - a datetime or date → the week containing that date
      - a string "YYYY-MM-DD" → parsed as a date

    If ref is a Sunday, it is treated as the END of that week (per your spec).
    """
    if ref is None:
        ref = datetime.now(LONDON).date()
    elif isinstance(ref, str):
        ref = datetime.strptime(ref, "%Y-%m-%d").date()
    elif isinstance(ref, datetime):
        ref = ref.date()

    # weekday(): Monday=0, Sunday=6
    days_since_monday = ref.weekday()  # 0 for Mon, 6 for Sun
    monday = ref - timedelta(days=days_since_monday)
    sunday = monday + timedelta(days=6)

    start = LONDON.localize(datetime(monday.year, monday.month, monday.day, 0, 0, 0))
    end   = LONDON.localize(datetime(sunday.year, sunday.month, sunday.day, 23, 59, 59))
    return start, end


def week_label(start: datetime, end: datetime) -> str:
    return f"{start.strftime('%d %b')} – {end.strftime('%d %b %Y')}"


def is_week_in_progress(end: datetime) -> bool:
    now = datetime.now(LONDON)
    return now < end
