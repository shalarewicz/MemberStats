# Utility
from dateutil import parser, tz
from sys import stderr
from datetime import timedelta
from tools import config
import argparse


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


def is_internal(address, interal_emails):
    """
    :return: True if the address contains 'irbnet.org' and is not listed
    in config.INTERNAL_EMAILS.
    """
    return 'irbnet.org' in address and all(x not in address for x in interal_emails)


def set_test(tst):
    """
    Assigns test values if the -t or --test flag (tools.TEST = True) was provided by the user
    :return:
    """
    if tst:
        # Testing Sheets / Ranges
        config.RETENTION_SPREADSHEET_ID = config.TEST_RETENTION_SPREADSHEET_ID
        config.ADMIN_SHEET = config.TEST_ADMIN_SHEET
        config.ADMIN_SHEET_ID = config.TEST_ADMIN_SHEET_ID
        config.MEMBER_STATS_SHEET = config.TEST_MEMBER_STATS_SHEET
        config.MEMBER_STATS_SHEET_ID = config.TEST_MEMBER_STATS_SHEET_ID
        config.WEEKLY_STATS_SPREADSHEET_ID = config.TEST_WEEKLY_STATS_SPREADSHEET_ID
        config.STATS_TO_ADDRESS = config.TEST_STATS_TO_ADDRESS
        config.ENROLLMENT_DASHBOARD_ID = config.TEST_ENROLLMENT_DASHBOARD_ID
        config.CURRENT_SHEET_ID = config.TEST_CURRENT_SHEET_ID
        config.GOV_MEMBER_STATS_SHEET = config.TEST_GOV_MEMBER_STATS_SHEET
        config.GOV_MEMBER_STATS_SHEET_ID = config.TEST_GOV_MEMBER_STATS_SHEET_ID
        config.GOV_WEEKLY_STATS_SHEET_ID = config.TEST_GOV_WEEKLY_STATS_SHEET_ID


def parse():
    """
    Parses and assigns global variables based on command line arguments.
    :return:
    """
    args = args_parser().parse_args()
    config.COUNT_ALL = args.all
    config.COUNT_EVERY = args.i
    config.COUNT_NONE = args.none
    config.SKIP = args.skip
    config.TEST = args.test
    config.DEBUG = args.debug
    config.GOV = args.gov
    # TODO if debug = True check that log directory exists


def args_parser():
    """
    Builds a parser with following attributes
     optional arguments:
    -h, --help            show this help message and exit
    -i I                  print every ith email read
    -a, --all             count all internal threads
    -n, --none            count no internal threads
    -t, --test 			  use test files and sheets not production sheets
    -g, --gov             only calculate government statistics
    -s, --skip            skips thread counting

    --auth_host_name AUTH_HOST_NAME
                          Hostname when running a local web server.
   --noauth_local_webserver
                         Do not run a local web server.
   --auth_host_port [AUTH_HOST_PORT [AUTH_HOST_PORT ...]]
                         Port web server should listen on.
   --logging_level {DEBUG,INFO,WARNING,ERROR,CRITICAL}
                         Set the logging level of detail.
    :return:
    """
    # Build args parser to validate mbox file-type and accept optional arguments
    arg_parser = argparse.ArgumentParser(description='Run weekly stats')

    # Optional: -i 1000 will print every 1000 email to help track progress
    arg_parser.add_argument("-i", type=int, default=0, help="print every ith email read")

    # Optional: -a | -n Automatically count or skip all internal emails.
    group = arg_parser.add_mutually_exclusive_group()
    group.add_argument("-a", "--all", action="store_true", help="count all internal threads")
    group.add_argument("-n", "--none", action="store_true", help="count no internal threads")
    group.add_argument("-s", "--skip", action="store_true",
                       help="skip thread counting. Use to test other parts of script")
    arg_parser.add_argument("-t", "--test", action="store_true", help="use test sheets not production sheets")
    arg_parser.add_argument("-d", "--debug", action="store_true", help="write mail and stat counting csv logs")
    arg_parser.add_argument("-g", "--gov", action="store_true", help="only calculate government statistics")

    return arg_parser


def print_param():
    """
    Prints the current operating parameters based on command line arguments.
    :return:
    """
    print "PARAMETERS:"
    if config.COUNT_ALL:
        print "    COUNT_ALL: All internal threads will be counted"
    if config.COUNT_NONE:
        print "    COUNT_NONE: No internal threads will be counted"
    if config.COUNT_EVERY > 0:
        print "    DISPLAY EVERY: Every " + str(config.COUNT_EVERY) + " emails"
    if config.SKIP:
        print "    SKIP: No threads will be counted. "
    if config.DEBUG:
        print "    DEBUG: Will write csv logs for mail and stat counting"
    if config.GOV:
        print "    GOV: Only government stats will be calculated. "
    if config.TEST:
        print "    TEST: Test sheets will be used rather than production sheets"
    else:
        print "    PRODUCTION: Production sheets will be used rather than test sheets"
