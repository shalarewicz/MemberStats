# Utility
from dateutil import parser, tz
from sys import stderr
from datetime import timedelta


def get_cutoff_date(message):
    """
    Prompt user to enter the earliest date for which stats should count.
    :return: A datetime object for the user entered date.
    """
    try:
        cutoff = parse_date(raw_input(message))
        if cutoff is None:
            print "Date not in an accepted format (MM/DD/YYYY)\nPlease try again."
            return get_cutoff_date(message)
        else:
            return cutoff
    except ValueError:
        print "Date not in an accepted format (MM/DD/YYYY)\nPlease try again."
        return get_cutoff_date(message)


def add_days(date, days):
    """
    Adds the specified number of days to date.
    :param date: original date
    :param days: days to be added.
    :return: the new date
    """
    return date + timedelta(days=days)


def parse_date(date, fuzzy=True, day_first=False, tz_info=tz.tzoffset('EDT', -14400)):
    """
    Parses the provided date and set the time zone to EDT if no time zone is specified.
    :param date: String date
    :param fuzzy: True if fuzzy parsing should be allowed (e.g. Today is January 1st(default=True).
    :param day_first: True if date should be parsed with the day first (e.g. DD/MM/YY) (default=False)
    :param tz_info: Default time zone info. (default = EDT GMT - 4)
    :return datetime object representing date.
    """
    try:
        ans = parser.parse(date, fuzzy=fuzzy, dayfirst=day_first)
        if ans.tzinfo is None:
            ans = ans.replace(tzinfo=tz_info)

        return ans
    except ValueError:
        # TODO Log this
        return None


# Return the A1 notation for a given index. _get_a1_column_notation(0) -> A
def get_a1_column_notation(i):
    """
    Converts an index i to the equivalent column in A1 notation. For example, A = 1 and AA = 27
    :param i: column index (A = 1)
    :return: A1 string notation
    """
    alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    j = i-1
    if j < 26:
        return alphabet[j]
    else:
        return alphabet[j / 26 - 1] + alphabet[j % 26]


def print_error(text):
    """
    Prints to standard error
    :param text: Error message
    :return: None
    """
    print >> stderr.write('\n'+str(text)+'\n')


def serial_date(date):
    """
    Converts date to a serial date as days since December 30th, 1899.
    :param date: datetime date
    :return: int Days since December 30th, 1899. Does not consider the time of day.
    """
    if date.tzinfo is None:
        date = date.replace(tzinfo=tz.tzoffset('EDT', -14400))
    zero = parse_date('12/30/1899')
    return (date - zero).days
