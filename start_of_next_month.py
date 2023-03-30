import datetime as dt


def start_of_next_month(date):
    """
    Returns a `datetime.date` object representing the first day of the next month after the specified date.
    Args:
        current_date (datetime.date): The date to use as a reference for calculating the next month.
    Returns:
        datetime.date: A `datetime.date` object representing the first day of the next month.
    """
    year = date.year + (date.month // 12)
    month = date.month % 12 + 1
    next_month = dt.date(year, month, 1)

    # Return the start of the next month
    return next_month