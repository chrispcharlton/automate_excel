"""
Contains helper functions. None are related to a specific Excel object.
"""

from datetime import datetime
from datetime import timedelta


def number_to_date(number: int) -> datetime.date:
    """Converts numbers to dates using datetime and timedelta"""
    date_origin = datetime(1899, 12, 30)
    new_date = date_origin + timedelta(days=number)
    return new_date


def date_to_number(date: datetime.date) -> int:
    """Converts dates to numbers using datetime"""
    date_origin = datetime(1899, 12, 30).date()
    number = (date - date_origin).days()
    return number


