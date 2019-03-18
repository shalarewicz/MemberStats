# Utility
from os.path import splitext
from dateutil import parser


def is_mbox(filename):
    """
    :param fname: filename
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
        cutoff = parser.parse(raw_input(
            "\nOn which date were stats last run?\n"
            "i.e. What is the earliest date for which stats should count (typically last Thursday)?\n"))
        return cutoff
    except ValueError:
        print "Date not in an accepted format (MM/DD/YYYY)\nPlease try again."
        return get_cutoff_date()