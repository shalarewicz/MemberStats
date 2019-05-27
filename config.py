import argparse
from oauth2client import tools  # TODO can probably remove this


def parser():
    # Build args parser to validate mbox file-type and accept optional arguments
    parse = argparse.ArgumentParser(description='Run weekly stats', parents=[tools.argparser])

    # Optional: -i 1000 will print every 1000 email to help track progress
    parse.add_argument("-i", type=int, default=0, help="print every ith email read")

    # Optional: -a | -n Automatically count or skip all internal emails.
    group = parse.add_mutually_exclusive_group()
    group.add_argument("-a", "--all", action="store_true", help="count all internal threads")
    group.add_argument("-n", "--none", action="store_true", help="count no internal threads")
    group.add_argument("-s", "--skip", action="store_true",
                       help="skip thread counting. Use to test other parts of script")
    parse.add_argument("-t", "--test", action="store_true", help="use test sheets not production sheets")

    return parse


def eval_test():
    global SPREADSHEET_ID, ADMIN_SHEET, MEMBER_STATS_SHEET, MEMBER_STATS_SHEET_ID, WEEKLY_STATS_SHEET_ID, STATS_EMAIL, \
        ENROLLMENT_DASHBOARD_ID, CURRENT_SHEET_ID
    # Testing Sheets / Ranges
    SPREADSHEET_ID = '1Vuozw7SwH4T-w6kAivL0nxLpdc3KpCFyQG8JeLlhPp8'
    ADMIN_SHEET = 'Support Outreach Administrators'
    MEMBER_STATS_SHEET = 'Member Stats'
    MEMBER_STATS_SHEET_ID = 1220379579
    WEEKLY_STATS_SHEET_ID = '1zT_lGeug1Nfk7x3RLmiT59Z3mVVBdv6ryqz-DRkh0q8'
    STATS_EMAIL = "stephan@irbnet.org"
    ENROLLMENT_DASHBOARD_ID = '1r454wPNgU9f1p8zc2BCCdytZ65A7SX1vq1QxdDbgutk'
    CURRENT_SHEET_ID = 1989883246


def print_param():
    print "PARAMETERS:"
    if COUNT_ALL:
        print "    COUNT_ALL: All internal threads will be counted"
    if COUNT_NONE:
        print "    COUNT_NONE: No internal threads will be counted"
    if COUNT_EVERY > 0:
        print "    DISPLAY EVERY: Every " + str(COUNT_EVERY) + " emails"
    if SKIP:
        print "    SKIP: No threads will be counted. "
    if TEST:
        print "    TEST: Test sheets will be used rather than production sheets"
    else:
        print "    PRODUCTION: Production sheets will be used rather than test sheets"


_args = parser().parse_args()
COUNT_ALL = _args.all
COUNT_EVERY = _args.i
COUNT_NONE = _args.none
SKIP = _args.skip
TEST = _args.test

# Production Sheets
SPREADSHEET_ID = '1mkxL43rqDyBZ6T8TIzg1_OQKhVNjefvYTDg9noC18j4'
ADMIN_SHEET = 'Support Outreach Administrators'
MEMBER_STATS_SHEET = 'Member Stats'
MEMBER_STATS_SHEET_ID = 1220379579
WEEKLY_STATS_SHEET_ID = '12wQxfv5EOEEsi3zCFwwwAq05SAgvzXoHRZbD33-TQ3o'
STATS_EMAIL = "andy@irbnet.org"
ENROLLMENT_DASHBOARD_ID = '1-NXL3_jHqH37-sfkVqSUS2cQ4XMUrILoyvt300iXjNQ'
CURRENT_SHEET_ID = 1989883246

eval_test()
print_param()