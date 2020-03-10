import argparse
from lib import util


# Production Sheets
SPREADSHEET_ID = '1mkxL43rqDyBZ6T8TIzg1_OQKhVNjefvYTDg9noC18j4'
ADMIN_SHEET = 'Support Outreach Administrators'
ADMIN_SHEET_ID = 452292035
MEMBER_STATS_SHEET = 'Member Stats'
MEMBER_STATS_SHEET_ID = 1220379579
WEEKLY_STATS_SHEET_ID = '12wQxfv5EOEEsi3zCFwwwAq05SAgvzXoHRZbD33-TQ3o'
STATS_EMAIL = "andy@irbnet.org"
ENROLLMENT_DASHBOARD_ID = '1g_EwipY4Yp1WXGrBhvw8Lly4fdieXSPaeiqTxnVGrpg'
CURRENT_SHEET_ID = 991688453


def set_test(tst):
    """
    Assigns test values if the -t or --test flag (tools.TEST = True) was provided by the user
    :return:
    """
    if tst:
        global SPREADSHEET_ID, ADMIN_SHEET, ADMIN_SHEET_ID, MEMBER_STATS_SHEET, MEMBER_STATS_SHEET_ID, \
            WEEKLY_STATS_SHEET_ID, STATS_EMAIL, ENROLLMENT_DASHBOARD_ID, CURRENT_SHEET_ID
        # Testing Sheets / Ranges
        SPREADSHEET_ID = '1Vuozw7SwH4T-w6kAivL0nxLpdc3KpCFyQG8JeLlhPp8'
        ADMIN_SHEET = 'Support Outreach Administrators'
        ADMIN_SHEET_ID = 452292035
        MEMBER_STATS_SHEET = 'Member Stats'
        MEMBER_STATS_SHEET_ID = 1220379579
        WEEKLY_STATS_SHEET_ID = '1zT_lGeug1Nfk7x3RLmiT59Z3mVVBdv6ryqz-DRkh0q8'
        STATS_EMAIL = "stephan@irbnet.org"
        ENROLLMENT_DASHBOARD_ID = '1r454wPNgU9f1p8zc2BCCdytZ65A7SX1vq1QxdDbgutk'
        CURRENT_SHEET_ID = 1989883246


# Support email settings. These addresses will be skipped when determining if a messages was sent internally.
INTERNAL_EMAILS = ["support@irbnet.org", "ideas@irbnet.org", "noreply@irbnet.org", "supportdesk@irbnet.org",
                   "techsupport@irbnet.org", "report_heartbeat@irbnet.org", "report_monitor@irbnet.org",
                   "alerts@irbnet.org", "wizards@irbnet.org", "reportmonitor2@irbnet.org"]

SUPPORT_EMAIL = "support@irbnet.org"
PING_EMAIL = "noreply@irbnet.org"
IDEAS_EMAIL = "ideas@irbnet.org"
SPAM_EMAILS = ["MAILER-DAEMON@LNAPL005.HPHC.org", "Mail Delivery System",
               "dmrn_exceptions@dmrn.dhhq.health.mil", "supportdesk@irbnet.org"]

# Thread "open" Labels. If any label contains any of the below phrases it will be considered open.
OPEN_LABELS = ["Waiting on", "TO DO", "To Call"]
VM_ADMIN= "vm/admin"
VM_SALES = "vm/sales"
VM_FINANCE = "vm/finance"
VM_RESEARCHER = "pings/vm"
PING_DEMO = "pings/demo"
PING_INQUIRY = "pings/inquiry"
PING_SUPPORT = "pings"
SALES_PING = "Sales Pings"
NEW_ORG = "New Organizations"
CHECK_IN = "check-in call"


# Command line args
COUNT_ALL = False
COUNT_EVERY = 0
COUNT_NONE = False
SKIP = False
TEST = False
DEBUG = False


def parse():
    """
    Parses and assigns global variables based on command line arguments.
    :return:
    """
    global COUNT_ALL, COUNT_EVERY, COUNT_NONE, SKIP, TEST, DEBUG
    args = parser().parse_args()
    COUNT_ALL = args.all
    COUNT_EVERY = args.i
    COUNT_NONE = args.none
    SKIP = args.skip
    TEST = args.test
    DEBUG = args.debug
    # TODO if debug = True check that log directory exists


def parser():
    """
    Builds a parser with following attributes
     optional arguments:
    -h, --help            show this help message and exit
    -i I                  print every ith email read
    -a, --all             count all internal threads
    -n, --none            skip all internal threads
    -t, --test 			  use test files and sheets not production sheets

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

    return arg_parser


def print_param():
    """
    Prints the current operating parameters based on command line arguments.
    :return:
    """
    print "PARAMETERS:"
    if COUNT_ALL:
        print "    COUNT_ALL: All internal threads will be counted"
    if COUNT_NONE:
        print "    COUNT_NONE: No internal threads will be counted"
    if COUNT_EVERY > 0:
        print "    DISPLAY EVERY: Every " + str(COUNT_EVERY) + " emails"
    if SKIP:
        print "    SKIP: No threads will be counted. "
    if DEBUG:
        print "    DEBUG: Will write csv logs for mail and stat counting"
    if TEST:
        print "    TEST: Test sheets will be used rather than production sheets"
    else:
        print "    PRODUCTION: Production sheets will be used rather than test sheets"


parse()
set_test(TEST)
print_param()

CUTOFF = util.get_cutoff_date("\nOn which date were stats last run?\n"
                              "i.e. What is the earliest date for which stats should count, typically last Thursday?\n")

# This date is inclusive of when stats should count.
END_CUTOFF = util.get_cutoff_date("\nEnter the final date for which stats should count. Typically this Wednesday.\n")
