# Utility
from os.path import splitext
from dateutil import parser, tz


def is_mbox(filename):
    """
    :param filename: filename
    :return: true if fname is an mbox file.
    """
    ext = splitext(filename)[1][1:]
    return ext == "mbox"


def update_date(old, new):
    if new > old:
        old = new
    return old


def get_cutoff_date():
    try:
        return parse_date(raw_input(
            "\nOn which date were stats last run?\n"
            "i.e. What is the earliest date for which stats should count (typically last Thursday)?\n"))
    except ValueError:
        print "Date not in an accepted format (MM/DD/YYYY)\nPlease try again."
        return get_cutoff_date()


def parse_date(date, fuzzy=True, day_first=False):
    try:
        ans = parser.parse(date, fuzzy=fuzzy, dayfirst=day_first)
        if ans.tzinfo is None:
            ans = ans.replace(tzinfo=tz.tzoffset('IST', 19800))

        return ans
    except ValueError:
        # TODO Log this
        return None


# Return the A1 notation for a given index. _get_a1_column_notation(0) -> A
def get_a1_column_notation(i):
    alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    j = i-1
    if j < 26:
        return alphabet[j]
    else:
        return alphabet[j / 26 - 1] + alphabet[j % 26]
