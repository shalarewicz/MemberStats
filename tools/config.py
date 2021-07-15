# Production Sheets
RETENTION_SPREADSHEET_ID = '1mkxL43rqDyBZ6T8TIzg1_OQKhVNjefvYTDg9noC18j4'
ADMIN_SHEET = 'Support Outreach Administrators'
ADMIN_SHEET_ID = 452292035
MEMBER_STATS_SHEET = 'Member Stats'
MEMBER_STATS_SHEET_ID = 1220379579
WEEKLY_STATS_SPREADSHEET_ID = '12wQxfv5EOEEsi3zCFwwwAq05SAgvzXoHRZbD33-TQ3o'
STATS_TO_ADDRESS = "andy@irbnet.org"
ENROLLMENT_DASHBOARD_ID = '1g_EwipY4Yp1WXGrBhvw8Lly4fdieXSPaeiqTxnVGrpg'
CURRENT_SHEET_ID = 991688453
RETENTION_CALLS = 'Weekly_Check_In'

# Gov Production Sheets
GOV_MEMBER_STATS_SHEET = 'Gov Member Stats'
GOV_MEMBER_STATS_SHEET_ID = 1102035090
GOV_WEEKLY_STATS_SHEET_ID = 1078993798

TEST_RETENTION_SPREADSHEET_ID = '16bXxrmnH6SC6DnUbgyLke-749GxBc4gw-6AhcHgXjZ0'
TEST_ADMIN_SHEET = 'Support Outreach Administrators'
TEST_ADMIN_SHEET_ID = 452292035
TEST_MEMBER_STATS_SHEET = 'Member Stats'
TEST_MEMBER_STATS_SHEET_ID = 1220379579
TEST_WEEKLY_STATS_SPREADSHEET_ID = '1zT_lGeug1Nfk7x3RLmiT59Z3mVVBdv6ryqz-DRkh0q8'
TEST_STATS_TO_ADDRESS = "stephan@irbnet.org"
TEST_ENROLLMENT_DASHBOARD_ID = '1r454wPNgU9f1p8zc2BCCdytZ65A7SX1vq1QxdDbgutk'
TEST_CURRENT_SHEET_ID = 1989883246
# Test Gov Sheets
TEST_GOV_MEMBER_STATS_SHEET = 'Gov Member Stats'
TEST_GOV_MEMBER_STATS_SHEET_ID = 1102035090
TEST_GOV_WEEKLY_STATS_SHEET_ID = 38016009

# Named Ranges
SHORT_NAME_RANGE = 'short_names'
GOV_SHORT_NAME_RANGE = 'gov_short_names'
ADMIN_CONTACT_INFO = 'admin_contact_info'

# Support email settings. These addresses will be skipped when determining if a messages was sent internally.
INTERNAL_EMAILS = ["support@irbnet.org", "ideas@irbnet.org", "noreply@irbnet.org", "supportdesk@irbnet.org",
                   "techsupport@irbnet.org", "report_heartbeat@irbnet.org", "report_monitor@irbnet.org",
                   "alerts@irbnet.org", "wizards@irbnet.org", "reportmonitor2@irbnet.org", "govsupport@irbnet.org"]

SUPPORT_EMAIL = "support@irbnet.org"
GOV_SUPPORT_EMAIL = "govsupport@irbnet.org"
PING_EMAIL = "noreply@irbnet.org"
IDEAS_EMAIL = "ideas@irbnet.org"
SPAM_EMAILS = ["MAILER-DAEMON@LNAPL005.HPHC.org", "Mail Delivery System",
               "dmrn_exceptions@dmrn.dhhq.health.mil", "supportdesk@irbnet.org"]

# Thread "open" Labels. If any label contains any of the below phrases it will be considered open.
OPEN_LABELS = ["Waiting on", "TO DO", "To Call"]
VM_ADMIN = "vm/admin"
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
COUNT_ALL = False  # If True, automatically counts all questioned threads
COUNT_EVERY = 0  # Prints every nth thread
COUNT_NONE = False  # If True, automatically counts no questioned threads
SKIP = False  # If True, skips thread counting
TEST = False  # IF True, uses test sheets
DEBUG = False  # IF True, outputs log files
GOV = False  # IF True, only calculates gov statistics

QUERY = " -label:no-reply -label:Report-Heartbeat -label:-googlespam -label:-180spam -label:WebEx " \
        "-label:-forwarded-to-govsupport -label:-spam"
