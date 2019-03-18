import csv
import mailbox
import sys
import time
import calendar
from recordtype import *
import argparse
import os.path
import traceback

import googleAPI
from members import *
from stats import *
from mail import *
from util import *
###########################################################################################
import httplib2
import os

# Construct a command line parser


def is_mbox_file(fname):
    """
    :param fname: filename
    :return: true if fname is an mbox file.
    """
    ext = os.path.splitext(fname)[1][1:]
    if not ext == "mbox":
        parser.error("Not a valid '.mbox' file")
    return fname


def parser():
    # Build args parser to validate mbox file-type and accept optional arguments
    parse = argparse.ArgumentParser(description='Run weekly stats', parents=[tools.argparser])

    # Required Argument: filename
    parse.add_argument("mbox_file", type=lambda s: is_mbox_file(s), help="mbox file for which stats will be gathered")

    # Optional: -i 1000 will print every 1000 email to help track progress
    parse.add_argument("-i", type=int, default=0, help="print every ith email read")

    # Optional: -a | -n Automatically count or skip all internal emails.
    group = parse.add_mutually_exclusive_group()
    group.add_argument("-a", "--all", action="store_true", help="count all internal threads")
    group.add_argument("-n", "--none", action="store_true", help="count no internal threads")
    group.add_argument("-s", "--skip", action="store_true", help="skip thread counting. Use to test other parts of script")
    parse.add_argument("-k", "--keep", action="store_true", help="keep mbox files when script has finished")
    parse.add_argument("-t", "--test", action="store_true", help="use test sheets not production sheets")


args = parser().parse_args()

# Store arguments for later use
MBOX = args.mbox_file
COUNT_ALL = args.all
COUNT_EVERY = args.i
COUNT_NONE = args.none
SKIP = args.skip
KEEP = args.keep
TEST = args.test 

print "PARAMETERS:"
print "    STATS FILE: " + MBOX
if COUNT_ALL:
    print "    COUNT_ALL: All internal threads will be counted"
if COUNT_NONE:
    print "    COUNT_NONE: No internal threads will be counted"
if COUNT_EVERY > 0:
    print "    DISPLAY EVERY: Every " + str(COUNT_EVERY) + " emails"
if SKIP:
    print "    SKIP: No threads will be counted. "
if KEEP:
    print "    KEEP: MBOX files will not be deleted"
if TEST:
    print "    TEST: Test sheets will be used rather than production sheets"
    # Testing Sheets / Ranges
    SPREADSHEET_ID = '1Vuozw7SwH4T-w6kAivL0nxLpdc3KpCFyQG8JeLlhPp8'
    ADMIN_SHEET = 'Support Outreach Administrators'
    MEMBER_STATS_SHEET = 'Member Stats'
    MEMBER_STATS_SHEET_ID = 1220379579
    WEEKLY_STATS_SHEET_ID = '1zT_lGeug1Nfk7x3RLmiT59Z3mVVBdv6ryqz-DRkh0q8'
    STATS_EMAIL = "stephan@irbnet.org"
    ENROLLMENT_DASHBOARD_ID = '1r454wPNgU9f1p8zc2BCCdytZ65A7SX1vq1QxdDbgutk'
    CURRENT_SHEET_ID = 1989883246
else:
    # Production Sheets / Ranges #TODO Update for 2019
    print "    PRODUCTION: Production sheets will be used rather than test sheets"
    SPREADSHEET_ID = '1mkxL43rqDyBZ6T8TIzg1_OQKhVNjefvYTDg9noC18j4'
    ADMIN_SHEET = 'Support Outreach Administrators'
    MEMBER_STATS_SHEET = 'Member Stats'
    MEMBER_STATS_SHEET_ID = 1220379579
    WEEKLY_STATS_SHEET_ID = '12wQxfv5EOEEsi3zCFwwwAq05SAgvzXoHRZbD33-TQ3o'
    STATS_EMAIL = "andy@irbnet.org"
    ENROLLMENT_DASHBOARD_ID = '1-NXL3_jHqH37-sfkVqSUS2cQ4XMUrILoyvt300iXjNQ'
    CURRENT_SHEET_ID = 1989883246

##################################################################################################

SHEETS_API = googleAPI.create_sheets_api()
MAIL_API = googleAPI.create_mail_api()


# Read in existing member stats
members = Member.read_members(MEMBER_STATS_SHEET, SPREADSHEET_ID, SHEETS_API)
print "\nThe following Statistics will be determined..."
for stat in sorted(STAT_LABELS): #todo does this know how to sort?
    print stat

Message.set_members(members.keys())

admins = Admin.read_admins(ADMIN_SHEET, SPREADSHEET_ID, SHEETS_API)

try:
    open_inquiries = OpenInquiry.from_file("Run Files/open.txt")
except IOError:
    print "ERROR: Could not read open.txt. Please ensure the file is formatted properly.\n", \
            "THREAD_ID\n," "SUBJECT"
    sys.exit()

cutoff = get_cutoff_date()

messages = {}
threads = {}

if not SKIP:
    print "\nPreparing " + MBOX
    i = 0
    for message in mailbox.mbox(MBOX):  # TODO Use Mail API to query support directly
        msg = Message(message)
        msg_id = msg.get_thread_id()
        messages[msg_id] = msg
        if msg_id not in threads and msg.counts:
            threads[msg_id] = Thread(msg)
        else:
            threads[msg_id].add_message(msg)

        # Todo Update admin dates. need to create a set of admin emails for lookup
        #  Use the last block of extractInfo()
        if "irbnet.org" not in message.get_from_address() and message.get_from_address() in admin_emails:
            pass

        if not COUNT_EVERY == 0 and i % COUNT_EVERY == 0:
            print i, msg
        i += 1

        # TODO Writers and lookup logs

    for thread in threads:
        if thread.get_oldest_date() < cutoff:
            # TODO Move to Thread class? make sure good thread doesn't flip back
            #  or check when counting stats to avoid going through all threads twice
            # If the thread has an oldest date before the cutoff change goodThread to false. This can
            # also be configured to ask the user if the thread should be counted. However, there are
            # a surprising amount of threads where this occurs 99% of which should not count so I
            # decided to automatically not count the thread. Therefore, if an admin replies to a thread
            # older than the cutoff with a completely new inquiry the thread won't be counted.
            # I'm OK with this as it does not happen too often.
            # TODO Figure out if the extract pulls all emails in the thread or just those within the
            #  specified timeframe
            thread.dont_count()

    # Add all pings and total labels to stats for counting.
    add_labels()

    new_open, new_closed = count_stats(threads)
    num_open, num_closed = OpenInquiry.update(open_inquiries, 'Inbox.mbox')

    # Update open inquiries
    STAT_LABELS["New Open Inquiries"].set_count(num_open)
    STAT_LABELS["Total Open Inquires"].set_count(num_open)
    STAT_LABELS["New Closed Inquiries"].set_count(num_closed)
    STAT_LABELS["Existing Open Inquiries Closed"].set_count(num_closed)
    STAT_LABELS["Total Closed Inquires"].set_count(new_closed + num_closed)

    # Combine and format in prep for writing
    format_stats()



# Read in existing member data and store
# Read in existing admin data and store
# Read in open inquiries and store

# Open the mbox file and parse each message
    # If this is the first message in the thread create a new thread
    # Else add the message to the thread AND the message is good THEN
        # update count by one
        # compare dates
        # add stat labels
        # add member labels
