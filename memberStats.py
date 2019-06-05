# name: memberStats.py
# author: Stephan Halarewicz
# date: 02/23/2018
#
# NEW IN THIS VERSION
# 	1. Version Control will now be through git (Repository: "https://github.com/stephanhalarewicz/MemberStats")
#	2. Integration of Google Sheets. Member retention and stats information is 
# 	   now be stored in 'Retention.gsheet'. Information will be written directly to this sheet
# 	   through use of the Google Sheets API. Data Support files ('mail.csv', 'formatted_mail.csv'
# 	   etc.) will still be stored in '.csv' format. These file will be stored in a directory 'Run Files'
# 	3. 'Weekly Support Stats.gsheet' will be updated automatically
#	4. TODO: Member stats will include a rolling average of how often they have contacted Support 
# 	   (weekly, monthly, quarterly, annually)
# 	5. Error handling to prevent data corruption
#	6. Python 3 will no longer be able to run this script as it is not supported by the Google Sheets API.
#	7. Use an argument parser
# 	8. Delete mbox files
#	9. Drafts and sends email to Andy
#
# REQUIREMENTS
# 	1. Python 2.7 
#  
#
# EXECUTE WITH: python memberStats filename.mbox
# 
# positional arguments:
#   mbox_file             mbox file for which stats will be gathered

# optional arguments:
#   -h, --help            show this help message and exit
#   -i I                  print every ith email read
#   -a, --all             count all internal threads
#   -n, --none            skip all internal threads
# 	-k, --keep            keep mbox files when script has finished
#  	-t, --test 			  use test files and sheets not production sheets

#   --auth_host_name AUTH_HOST_NAME
#                         Hostname when running a local web server.
#   --noauth_local_webserver
#                         Do not run a local web server.
#   --auth_host_port [AUTH_HOST_PORT [AUTH_HOST_PORT ...]]
#                         Port web server should listen on.
#   --logging_level {DEBUG,INFO,WARNING,ERROR,CRITICAL}
#                         Set the logging level of detail.

"""
    Read Institution short names from Retention tab
        add to members dict
    Read member stats from member info tab
        read and store labels
    Read admins from support admins
        store in admins dict (name, adminInfo)
    Update admin contact dates
        use a named range and stored row # in admin info
        used stored information to update sheet

"""


import csv
import mailbox
import sys
import time
import calendar
from recordtype import *
import argparse
import os.path
import traceback


###########################################################################################
# Imports for Google Client Authentication
import os

from tools.lib.googleAPI import MAIL_API, SHEETS_API
from oauth2client import tools
from googleapiclient.errors import HttpError
from email.mime.text import MIMEText
import base64

###########################################################################################

# isMboxFile() str --> boolean. Checks if the given file is an mbox file
def isMboxFile(fname):
    ext = os.path.splitext(fname)[1][1:]
    if not ext == "mbox":
        parser.error("Not a valid '.mbox' file")
    return fname

# Build args parser to validate mbox file-type and except optional arguments
parser = argparse.ArgumentParser(description='Run weekly stats', parents=[tools.argparser])

# Required. Check for mbox filetype
parser.add_argument("mbox_file", type=lambda s:isMboxFile(s), 
    help="mbox file for which stats will be gathered")
# Optional: -i 1000 will print every 1000 email to help track progress
parser.add_argument("-i", type=int, default=0, help="print every ith email read")

# Optional: -a | -n Automatically count or skip all internal emails. 
group = parser.add_mutually_exclusive_group()
group.add_argument("-a", "--all", action="store_true", help="count all internal threads")
group.add_argument("-n", "--none", action="store_true", help="count no internal threads")
group.add_argument("-s", "--skip", action="store_true", help="skip thread counting. Use to test other parts of script")
parser.add_argument("-k", "--keep", action="store_true", help="keep mbox files when script has finished")
parser.add_argument("-t", "--test", action="store_true", help="use test sheets not production sheets")
args = parser.parse_args()

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
    print "    DISPLAY EVERY: Every " + str(COUNT_EVERY) +" emails"
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
    # Production Sheets / Ranges
    print "    PRODUCTION: Production sheets will be used rather than test sheets"
    SPREADSHEET_ID = '1mkxL43rqDyBZ6T8TIzg1_OQKhVNjefvYTDg9noC18j4'
    ADMIN_SHEET = 'Support Outreach Administrators'
    MEMBER_STATS_SHEET = 'Member Stats'
    MEMBER_STATS_SHEET_ID = 1220379579
    WEEKLY_STATS_SHEET_ID = '12wQxfv5EOEEsi3zCFwwwAq05SAgvzXoHRZbD33-TQ3o'
    STATS_EMAIL = "andy@irbnet.org"
    ENROLLMENT_DASHBOARD_ID = '1g_EwipY4Yp1WXGrBhvw8Lly4fdieXSPaeiqTxnVGrpg'
    CURRENT_SHEET_ID = 991688453

##################################################################################################
# B. Read in existing member info and stats to be counted

# 1. Get stats to be counted
# 2. Read in existing information or can I just add to existing value later
# 3. Assign priority to each stat

# Returns values in the provided range of the Retention sheet.
def getRange(range):
    return SHEETS_API.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID,
        range=range, majorDimension='ROWS').execute().get('values',[])

members = {}
memberInfo = recordtype('memberInfo', 'lastContact phone stats')

statsLabels	= {}
statInfo = recordtype('statInfo', 'count priority', default = 0)

# Read the entire sheet
memberStats = getRange(MEMBER_STATS_SHEET)

# Read stats to be counted
colnum = 0
priority = 0
for col in memberStats[0]:
    if colnum == 0 or colnum == 1 or colnum == 2:
        colnum += 1
        continue
    else:
        stat = memberStats[0][colnum]
        statsLabels[stat] = statInfo(0, priority)
        priority +=1
        colnum += 1

NUMBER_OF_MEMBER_STATS = len(statsLabels)

# Read in existing member stats
onHeader = True
rowNum = 0
for row in memberStats:
    if onHeader:
        onHeader = False
        continue
    else:
        colnum = 0
        mem = ""
        date = ""
        phone = ""
        stats = []
        for col in row:
            if colnum == 0:
                mem = row[colnum]
            elif colnum == 1:
                date = row[colnum]
            elif colnum == 2:
                phone = row[colnum]
            else:
                stats.append( int(row[colnum]) )
            colnum += 1

        members[mem] = memberInfo(date, phone, stats)
        rowNum += 1

# Read in member short names on Retention sheet to ensure all members accounted for.
memberListRange = 'Short_Names' # Named Range of members
memberList = getRange(memberListRange)

emptyStatRow = [0] * NUMBER_OF_MEMBER_STATS
for mem in memberList:
    if mem[0] not in members.keys():
        print mem
        members[mem[0]] = memberInfo("", "", emptyStatRow)



statsLabels["Sales Pings"] = statInfo(0, priority + 30)

print "\nThe following Statistics will be determined..."
for stat in sorted(statsLabels):
    print stat

#################################################################################################
# C. Read in Administrator Information

# Dictionary for storing administrator data key = name, value = recordtype(adminInfo)
admins = {}
adminEmails = {} #not sure if this is needed key = email value = name used to cross-reference data

# Keep track of emails which aren't tracked. 
missedAdmins = {} 

#Record Types for storing admin info
adminInfo = recordtype('adminInfo', 'row checkIn lastContact')

missedAdminsFile = open('Run Files\Admin Contacts Not Tracked.csv', 'rb')
missedAdminReader = csv.reader(missedAdminsFile)
missedAdminHeader = []

onHeader = True
# Read in existing admin email addresses which are not tracked for writing later
for adminEmail in missedAdminReader:
    if onHeader:
        missedAdminHeader = adminEmail
        onHeader = False
        continue
    missedAdmins[adminEmail[0]] = adminEmail[1]

missedAdminsFile.close()

#Read in existing admin info
onHeader = True

adminHeader = []

adminSheet = getRange(ADMIN_SHEET)

rownum = 0
for admin in adminSheet:
    if onHeader:
        onHeader = False
        continue
    else:
        try:
            name = admin[1]
            checkInDate = admin[4]
            lastContactDate = admin[5]
            email1 = admin[6].lower()
            email2 = admin[7].lower()
            email3 = admin[8].lower()

            admins[name] = adminInfo(rownum, checkInDate, lastContactDate)
            adminEmails[email1] = name
            #Need to check other emails as well
            if not email2 == "":
                adminEmails[email2] = name
            if not email3 == "":
                adminEmails[email3] = name
            rownum +=1
        except IndexError:
            print "\n\nCould not read full data from Support Outreach Administrators tab for:"
            print "\t" + str(admin) + "\n"
            print "This is likely the result of a missing phone number. Please ensure the following fields are entered for this administrator and re-run the script."
            print "Institution"
            print "Name"
            print "Role"
            print "Department			"
            print "Primary E-mail"
            print "Primary Phone"

            sys.exit()


#######################################################################################
# C.  Read in open inquiries
# Read in current open inquiries from open.txt. 
#
# 	open.txt should have format:
#		threadID
#	 	subject
#	 	"Open" or Closed"
#		"Y" or "N" - indicates if thread should be counted. 
# 
# A text file is used to store this information rather than a csv due to Excel truncating thread ID since
# the number is to big. 

# openInquiries used to store information
openInquires = {}
openInfo = recordtype('openInfo', 'subject, open, good')

lastWeekOpen = 0

print "\nBegin\n"

# Open open.txt for reading
openInquiriesFile = open('Run Files/open.txt', 'rb')

currentID = ""
currentSubject = ""
currentOpen = False
currentGood	 = False
i = 1
print "Reading open.txt..."
for line in openInquiriesFile:
    if i == 1:
        currentID = line.strip()
        i = 2
        continue
    elif i == 2:
        currentSubject = line.strip()
        i = 3
        continue
    elif i == 3:
        if line.strip() == "Open":
            currentOpen = True
        if line.strip() == "Closed":
            currentOpen = False
        i = 4
        continue
    elif i == 4:
        if line.strip() == "Y":
            currentGood = True
        if line.strip() == "N":
            currentGood = False
        i = 1
        openInquires[currentID] = openInfo(currentSubject, currentOpen, currentGood)
        lastWeekOpen += 1
        continue
    else:
        print "Invalid i value"

openInquiriesFile.close()

###########################################################################################################

# D. Helper Functions
# 1. The call isMember(str label) --> boolean returns true if label is in members
def isMember(label):
    return label in members

# 2. The call isStat(str label) --> boolean returns true if label is in statsLabels
def isStat(label):
    return label in statsLabels

# 3. The call findLabels(str labels) --> str list returns a list of labels contained 
#    within the given dictionary from a string of comma-separated labels
#
# 	 findLabels(members, Tulane,ghc,Information) --> [ghc, Tulane]
# 	 findLabels(stats, Tulane,ghc,Information) --> [Information]

def findLabels(dictionary, labels):
    split = labels.split(",")
    result = []
    for label in split:
        if label in dictionary:
            result.append(label)
    return result

# 4. The call isSpam(str frm) --> boolean checks to see if the from address contains 
# 	 any of the following phrases or addresses. This is used to save time and filter out 
# 	 Spam and delivery failures. Note that other addresses can be added if desired. 

def isSpam(frm):
    if "<MAILER-DAEMON@LNAPL005.HPHC.org>" in frm:
        return True
    elif "Mail Delivery System" in frm:
        return True
    elif "dmrn_exceptions@dmrn.dhhq.health.mil" in frm:
        return True
    elif "<supportdesk@irbnet.org>" in frm:
        return True
    elif "" == frm:
        return True
    else:
        return False

# 5. The call isIdea(str frm) --> boolean checks to see if the from address matches 
# 	 'ideas@irbnet.org'. This is used to save time and filter out Produce Enhancement 
# 	 emails

def isIdea(two):
    return "<ideas@irbnet.org>" in two

# 6. The call checkToFromSupport(str two, str frm) --> boolean checks to see if a 
# 	 message was sent to Support from Support. This is used to help handle cases
# 	 when an admin may have started a new thread. 

def checkToFromSupport(two, frm):
    return "support@irbnet.org" in two and "support@irbnet.org" in frm

# 7. The call checkFromSupport(str frm) --> boolean returns True if the email is 
#  	 from Support. This is used to determine if a date should be used for the date of
# 	 last contact with a member. 

def checkFromSupport(frm):
    return "support@irbnet.org" in frm

# 8. The call isInternal(str frm) checks to see if an email was sent by someone within
#	 IRBNet. This is used to help avoid counting internal threads. 

def isInternal(frm):
    return "irbnet.org" in frm and "support@irbnet.org" not in frm and "ideas@irbnet.org" not in frm and "noreply@irbnet.org" not in frm and "supportdesk@irbnet.org" not in frm and "techsupport@irbnet.org" not in frm and "report_heartbeat@irbnet.org" not in frm and	"report_monitor@irbnet.org" not in frm and "alerts@irbnet.org" not in frm and "wizards@irbnet.org" not in frm and "reportmonitor2@irbnet.org" not in frm

# 9. createDate(str day, str month, str year) --> str Converts date into a usable 
# 	 format MM/DD/YYYY.
#    the call createDate("25", Apr, "1994") will result in "04/25/1994"

def createDate(day, month, year):
    if month == "Jan" or month == "01" or month == "January":
        newMonth = "01"
    elif month == "Feb" or month == "02" or month == "February":
        newMonth = "02"
    elif month == "Mar" or month == "03" or month == "March":
        newMonth = "03"
    elif month == "Apr" or month == "04" or month == "April":
        newMonth = "04"
    elif month == "May" or month == "05" or month == "May":
        newMonth = "05"
    elif month == "Jun" or month == "06" or month == "June":
        newMonth = "06"
    elif month == "Jul" or month == "07" or month == "July":
        newMonth = "07"
    elif month == "Aug" or month == "08" or month == "August":
        newMonth = "08"
    elif month == "Sep" or month == "09" or month == "September":
        newMonth = "09"
    elif month == "Oct" or month == "10" or month == "October":
        newMonth = "10"
    elif month == "Nov" or month == "11" or month == "November":
        newMonth = "11"
    elif month == "Dec" or month == "12" or month == "December":
        newMonth = "12"
    else:
        print "The three letter month is not valid..." + month
    return newMonth + "/" + day + "/" + year

# 10.compareDates(str date1, str date2) --> boolean compares two dates of format 
# 	 MM/DD/YYYY and returns true if date1 is more recent or equal to date2
#	 The call compareDates("04/25/1994", "02/17/2017") returns False

def compareDates(date1, date2):
    if date2 == "":
        return True;
    if date1 == "":
        return False
    info1 = date1.split("/")
    info2 = date2.split("/")
    if int(info1[2]) > int(info2[2]):
        return True
    elif	int(info1[2]) < int(info2[2]):
        return False
    elif	int(info1[0]) > int(info2[0]):
        return True
    elif	int(info1[0]) < int(info2[0]):
        return False
    elif	int(info1[1]) > int(info2[1]):
        return True
    elif	int(info1[1]) < int(info2[1]):
        return False
    else:
        return True

# 11.prepDate(str date) --> str
#    prepDate takes a date string and returns a formatted date if the original
# 	 date string is acceptable. Note that this method is specific to GMail Date Strings
# 	 If date string not acceptable method will return "00/00/0000"
# 	 Commented code is useful for troubleshooting

def prepDate(date):
    try:
        splits = date.split(",")
        if len(splits) < 2:
            # print "Bad Date..." + date
            dates = splits[0].split(" ")
            date = createDate(dates[0], dates[1], dates[2])
            return date
        # print "Converting..." + date
        # print splits
        # print row
        dates = splits[1].split(" ")
        # print dates
        if len(dates) > 4:
            if dates[1] == "":
                date = createDate(dates[2], dates[3], dates[4])
                return date
            else:
                date = createDate(dates[1], dates[2], dates[3])
                return date
        else:
            print "Bad Date...." + date
            return "00/00/0000"
    except IndexError:
        print "IndexError for the following date."
        print date
        return "00/00/0000"


# 12. extractInfo( str threadID, str labels, str date, boolean goodMessage) --> void
# 	  The call extractInfo takes the given information, extracts all member and stats 
# 	  labels and formats the date using prepDate. If the threadID already exists in 
# 	  the dictionary threads then the date is compared against the oldest date recorded 
# 	  for the thread and the last contact date for the thread and replaces them as 
# 	  necessary. Note that the last contact date will only be updated if goodMessage is 
# 	  passed as true (goodMessage is only false when an email is from Support). If the 
# 	  the thread does not exist then a new entry in threads is created for key threadID.

# Record type used to organize info in dictionary 'threads'
threadInfo = recordtype('threadInfo', 
    'statsLabels memberLabels goodThread date oldestDate nonPing demo inquiry vm count checked closed checkIn checkInDate')

def extractInfo(threadID, labels, date, goodMessage):
    threadStats = findLabels(statsLabels, labels)
    threadMembers = findLabels(members, labels)
    newDate = prepDate(date)
    if threadID in threads:
        for stat in threadStats:
            if stat not in threads[threadID].statsLabels:
                threads[threadID].statsLabels.append(stat)
        for mem in threadMembers:
            if mem not in threads[threadID].memberLabels:
                threads[threadID].memberLabels.append(mem)
        if compareDates(newDate, threads[threadID].date) and goodMessage:
            threads[threadID].date = newDate
        if compareDates(threads[threadID].oldestDate, newDate):
            threads[threadID].oldestDate = newDate
    elif goodMessage:
        threads[threadID] = threadInfo(threadStats, threadMembers, True,
            newDate, newDate, False, False, False, False, 0, False, True, False, "")
    else:
        threads[threadID] = threadInfo(threadStats, threadMembers, True,
            "", newDate, False, False, False, False, 0, False, True, False, "")
    if "check-in call" in labels:
        threads[threadID].checkIn = True
        if compareDates(newDate, threads[threadID].checkInDate):
            threads[threadID].checkInDate = newDate
    if threads[threadID].closed and (("Waiting on" in labels) or ("TO DO" in labels) or ("To Call" in labels)):
        threads[threadID].closed = False
    # Compares message date against dates for admins
    if not "@irbnet.org" in fromEmail:
        if fromEmail in adminEmails:
            adminContactDate = admins[adminEmails[fromEmail]].lastContact
            adminCheckInDate = admins[adminEmails[fromEmail]].checkIn
            if compareDates(newDate, adminContactDate):
                admins[adminEmails[fromEmail]].lastContact = newDate
            if threads[threadID].checkIn and compareDates(newDate, adminCheckInDate):
                admins[adminEmails[fromEmail]].checkIn = newDate



# 13. shouldItCount(str fromAddress, str to, str subject, str date, str labels, str emailType) --> void
#     Ask the user if an email should be counted or not. The call will print relevant email 
# 	  information and ask the user to respond with 'Y' or 'N'. If the users responds with no then
# 	  the thread will be marked as bad (goodThread = bad) and not be counted. If the user enters
# 	  an unrecognized response the user will be prompted to answer again.


def shouldItCount(fromAddress, to, subject, date, labels, emailType):
    if not COUNT_ALL or COUNT_NONE:
        threads[threadID].checked = True
        print ""
        print "Found " + emailType +" email. Should the following message be counted?"
        print ""
        print "From: " + fromAddress
        print "To: " + to
        print "Subject: " + subject
        print "Date: " + date
        print "Labels:" + labels
        answer = raw_input("Y/N?    ")
        if answer == "Y" or answer == "y":
            print "Thread will be counted."
            threads[threadID].goodThread = True
        elif answer == "N" or answer == "n":
            print "Thread won't be counted."
            threads[threadID].goodThread = False
        else:
            print "Answer not recognized."
            shouldItCount(fromAddress, to, subject, date, labels, emailType)
    elif COUNT_NONE:
        threads[threadID].goodThread = False
    else:
        threads[threadID].goodThread = True

# 14. The call writeInfo(str threadID, str stat) writes basic thread information for
# 	  the given threadID along with the stat that was counted. Used to write information
# 	  to threadLookupInfo.csv

def writeInfo(threadID, stat, closed):
    lookupWriter.writerow([threadLookupInfo[threadID][0], threadLookupInfo[threadID][1],
        threadLookupInfo[threadID][2], threadLookupInfo[threadID][3], threadLookupInfo[threadID][4],
        stat, closed])

# 15. The call sortStats(str list stats) -> str list returns the given list of stats
#  	  sorted by the stats assigned priority. Sorted lo --> hi priority

def sortStats(stats):
    less = []
    equal = []
    greater = []
    if len(stats) > 1:
        pivot = statsLabels[stats[0]].priority
        for stat in stats:
            if statsLabels[stat].priority < pivot:
                less.append(stat)
            if statsLabels[stat].priority == pivot:
                equal.append(stat)
            if statsLabels[stat].priority > pivot:
                greater.append(stat)
        return sortStats(less) + equal + sortStats(greater)
    else:
        return stats

# 16. The call extractEmail(email) will return a string containing the email address from any string
#	  input containing an email address between a '<' and '>'

def extractEmail(email):
    if '<' not in email or '>' not in email:
        return email
    start = email.find('<')
    end = email.find('>')
    start += 1
    result = ""
    while start < end:
        result = result + email[start]
        start += 1
    return result.lower()

# 17. The call sortedAdmins(dict<key, value> = <string, recordtype(adminInfo>)) will return a list of
#	  keys sorted alphabetically by adminInfo.org

def sortedAdmins(adminDictionary):
    copy = adminDictionary
    less = {}
    equal = {}
    equalList = []
    greater = {}

    findPivot = True

    if len(copy) >= 1:
        for admin in copy:
            if findPivot:
                pivot = copy[admin].org
                equalList.append(admin)
                findPivot = False
                continue
            if copy[admin].org < pivot:
                less[admin] = copy[admin]
            if copy[admin].org == pivot:
                equal[admin] = copy[admin]
                equalList.append(admin)
            if copy[admin].org > pivot:
                greater[admin] = copy[admin]
        return sortedAdmins(less) + equalList + sortedAdmins(greater)
    else:
        return []


###############################################################################################

# E. Now for the fun part

# 1. Open mail.csv and formatted_mail.csv for writing then write the headers

firstOutFile = open('Run Files\mail.csv', 'wb')

writer = csv.writer(firstOutFile)

outfile = open('Run Files\\formatted_mail.csv', "wb")
datedWriter = csv.writer(outfile)

writer.writerow(['Thread ID', 'Date', 'From', 'To', 'Subject', 'X-Gmail-Labels'])
datedWriter.writerow(['Thread ID', 'Date', 'From', 'To', 'Subject', 'Stats', 'Members'])

# i = message counter
i = 0

# Keep track of thread info
threads = {}

# For printing messages to threadLookupInfo.csv
threadLookupInfo = {}

# Counters for various stats
pingInquiries = 0
pingDemos = 0
voicemails = 0
newOrgs = 0
salesPings = 0
webForm = 0

# User will be asked to enter the earlies date for which a thread should be counted. Format should be
# MM/DD/YYYY

def getCutoffDate():
    cutoff = raw_input("\nOn which date were stats last run?\ni.e. What is the earliest date for which stats should count\n(MM/DD/YYYY)    ")
    try:
        time.strptime(cutoff, "%m/%d/%Y")
        return cutoff
    except:
        print "Date not in correct format (MM/DD/YYYY)\nPlease try again"
        return getCutoffDate()

cutoff = getCutoffDate()

if not SKIP:
    print "\nPreparing " + MBOX
    for message in mailbox.mbox(MBOX):
        if i == 0:
            print "\nStarting to read messages"
            i += 1

        threadID = message['X-GM-THRID']
        labels = message['X-Gmail-Labels']
        date = message['Date']
        fromAddress = message['From']
        to = message['To']
        subject = message['Subject']

        # If any of the field is blank convert from None to "" to avoid errors
        if labels is None:
            writer.writerow([threadID, message['Date'], message['From'], message['To'],
             message['Subject'], message['X-Gmail-Labels']])
            continue
        if threadID is None:
            threadID = ""
        if fromAddress is None:
            fromAddress = ""
        if to is None:
            to = ""
        if subject is None:
            subject = ""
        if date is None:
            date = ""
        # Replace newline characters ('\r' and '\n') to avoid separating labels with spaces. This is a result of
        # a long label string in the mbox file.
        if "\r" or "\n" in labels:
            labels = labels.replace('\n', '').replace('\r', '')
        if "\r" or "\n" in subject:
            subject = subject.replace('\n', '').replace('\r', '')

        # If spam or idea then write to mail.csv then continue. No need to waste time formatting.
        if isSpam(fromAddress):
            writer.writerow([threadID, message['Date'], message['From'], message['To'],
             message['Subject'], message['X-Gmail-Labels']])
            i += 1
            continue

        if isIdea(to):
            writer.writerow([threadID, message['Date'], message['From'], message['To'],
                message['Subject'], message['X-Gmail-Labels']])
            i += 1
            continue

        # Extract email from fromAddress currently formatted as "Last name, First name" <email>
        fromEmail = extractEmail(fromAddress)

        # Check for emails from Support. If email is from Support then pass False to extractInfo
        # to not use the email date for lastContact and vice versa
        try:
            if checkFromSupport(fromAddress):
                extractInfo(threadID, labels, date, False)
            else:
                extractInfo(threadID, labels, date, True)
        except IndexError:
            print "IndexError for the following date."
            print message['Date']
            continue

        # If the the thread contains a stat labels then change nonPing to True
        if not threads[threadID].statsLabels == []:
            threads[threadID].nonPing = True

        if "@irbnet.org" not in fromEmail and threads[threadID].nonPing and fromEmail not in adminEmails and fromEmail not in missedAdmins:
            missedAdmins[fromEmail] = ""

        # If the thread has an oldest date before the cutoff change goodThread to false and write
        # message to mail.csv then continue. This can also be configured to ask the user if the
        # thread should be counted however there are a surprising amount of threads that have this
        # problem 99% of which should not count so I decided to skip this. Therefore, if an admin
        # replies to a thread from a few months ago with a completely new inquiry the thread won't
        # be counted. I'm OK with this as it does not happen too often.

        if not compareDates(threads[threadID].oldestDate, cutoff):
            threads[threadID].count += 1
            threads[threadID].goodThread = False
            writer.writerow([threadID, message['Date'], message['From'], message['To'],
                message['Subject'], message['X-Gmail-Labels']])
            continue

        # If thread is bad write the message to mail.csv and continue
        if not threads[threadID].goodThread:
            threads[threadID].count += 1
            writer.writerow([threadID, message['Date'], message['From'], message['To'],
             message['Subject'], message['X-Gmail-Labels']])
            i += 1
            continue

        # If threadID in threads then if the thread is a ping or new org and the thread count = 2 change
        # either inquiry demo vm to True or increase newOrg counter by one (change newOrg later)
        #
        # If the thread is nonPing and the current message is to and from Support or sent internally
        # ask if the thread should count
        #
        # If 'Sales Ping is in the labels for the message' and it was sent via webform increase the appropriate
        # counter.

        threads[threadID].count += 1
        if not threads[threadID].nonPing:
            if "IRBNet Demo Request" in subject and threads[threadID].count == 2:
                    threads[threadID].demo = True
                    threads[threadID].goodThread = True
            if "IRBNet Inquiry From" in subject and threads[threadID].count == 2:
                threads[threadID].inquiry = True
                threads[threadID].goodThread = True
            if "IRBNet Help Desk Inquiry" in subject:
                if "noreply@irbnet.org" not in fromAddress:
                    threads[threadID].count -= 1
                threads[threadID].vm = True
                threads[threadID].goodThread = True
            if "New Organizations" in labels and threads[threadID].count == 2:
                newOrgs += 1
                threads[threadID].goodThread = True
        else:
            if checkToFromSupport(to, fromAddress) and not "New Organizations" in threads[threadID].statsLabels and not threads[threadID].checked:
                shouldItCount(fromAddress, to, subject, date, labels, "to and from Support")
            elif (isInternal(fromAddress) or (checkFromSupport(fromAddress) and isInternal(to) and not "Sales Pings" in threads[threadID].statsLabels)) and not threads[threadID].checked:
                shouldItCount(fromAddress, to, subject, date, labels, "Internal")
            if "Sales Ping" in labels and ("IRBNet Demo Request" in subject or "IRBNet Inquiry From" in subject) and not threads[threadID].checked:
                    webForm += 1
                    threads[threadID].checked = True
                    threads[threadID].goodThread = True


        # Write message to formatted_mail.csv
        datedWriter.writerow([threadID, threads[threadID].date, message['From'], message['To'],
         message['Subject'], threads[threadID].statsLabels, threads[threadID].memberLabels])

        # Write message to mail.csv
        writer.writerow([threadID, message['Date'], message['From'], message['To'],
         message['Subject'], message['X-Gmail-Labels']])


        # Print every user specified amount of messages.
        if not COUNT_EVERY == 0 and i % COUNT_EVERY == 0:
            print i
            print [threadID, message['Date'], message['From'], message['To'],
         message['Subject'], message['X-Gmail-Labels']]

        # Increase message count by one
        i += 1

        # Add message info to the dictionary 'threadLookupInfo' to store data when counting stats and
        # writing threadLookupInfo.csv

        if not threadID in threadLookupInfo:
            threadLookupInfo[threadID] = [threads[threadID].oldestDate, subject, fromAddress, to, labels]

    # Close mail.csv
    # Close formatted_mail.csv
    firstOutFile.close()
    outfile.close()

    print "\nDone reading mbox file\n"

    #######################################################################################
    # F. Update date of last contact and check-in date for each member

    # 1. compare the most recent date for each thread against the date of last contact
    #    for each member and swap the them in necessary.

    print "Updating dates of last contact and check-in call dates...\n"

    for thread in threads:
        if threads[thread].memberLabels == []:
            pass
        else:
            for mem in threads[thread].memberLabels:
                if compareDates(threads[thread].date, members[mem].lastContact):
                    members[mem].lastContact = threads[thread].date
                if compareDates(threads[thread].checkInDate, members[mem].phone):
                    members[mem].phone = threads[thread].checkInDate

#######################################################################################
# G. Count Statistics

# Open 'threadLookupInfo.csv' for writing
# This is done to cross reference which stat was counted for a thread if there is ever a dispute


lookupFile = open('Run Files\\threadLookupInfo.csv', 'wb')
lookupWriter = csv.writer(lookupFile)

# Write headers. 
lookupWriter.writerow(["Date", "Subject", "From", "To", "Labels", "Counted Stat", "Closed?"])


# For every good thread if their are no stats labels check to see if the thread is a true ping
# A true ping is one that has nonPing False and either inquiry, demo, or vm True. This appears to 
# be redundant however remember that if an admin replies to a ping the thread will have both nonPing
# and ping marked as true. 
#
# If the stat list isn't empty then sort the list and use the first list item to count for stats.
# If a non-ping is open add it to openInquiries and increase newOpen by one.  
# Increase the count for that stat by one then write info to threadLookupInfo.csv. 
#
# For each member in memberLabels increase the member specific stat count for each stat in statLabels
# Note: members[mem].stats can be indexed by the stat priority due to the way the stats list is 
# constructed and the stat priority is assigned. (priority starts at 0)

newClosed = 0
newOpen = 0

for thread in threads:
    if threads[thread].goodThread:
        if threads[thread].closed:
            threadClosed = "Closed"
        else:
            threadClosed = "Open"
        if threads[thread].statsLabels == []:
            if not threads[thread].nonPing and threads[thread].inquiry:
                pingInquiries += 1
                writeInfo(thread, "Inquiry", threadClosed)
            if not threads[thread].nonPing and threads[thread].demo:
                pingDemos += 1
                writeInfo(thread, "Demo", threadClosed)
            if not threads[thread].nonPing and threads[thread].vm:
                writeInfo(thread, "Voicemail", threadClosed)
                voicemails += threads[thread].count
        else:
            sortedStats = sortStats(threads[thread].statsLabels)
            if sortedStats is None:
                # This was originally for trouble shooting but don't think its necessary anymore.
                print "Found sortedStats = None during counting..."
                print threadLookupInfo[thread][0]
                print threadLookupInfo[thread][1]
                print threadLookupInfo[thread][2]
                print threadLookupInfo[thread][3]
                print threadLookupInfo[thread][4]
            else:
                toCount = sortedStats[0]
                statsLabels[toCount].count += 1
                # Sales Pings and Sales are considered Pings
                if threads[thread].closed and (not toCount == "Sales Pings" and not toCount == "Sales"):
                    newClosed += 1
                    writeInfo(thread, toCount, threadClosed)
                elif not threads[thread].closed and (not toCount == "Sales Pings" and not toCount == "Sales"):
                    newOpen += 1
                    openInquires[thread] = openInfo(threadLookupInfo[thread][1], True, True)
                    writeInfo(thread, toCount, threadClosed)
                try:
                    for mem in threads[thread].memberLabels:
                        for stat in sortedStats:
                            statPriority = statsLabels[stat].priority
                            members[mem].stats[statPriority] += 1
                except IndexError:
                    print "Mem Specific index error"
                    print members[mem].stats
                    print sortedStats
                    sys.exit()

# Close threadLookupInfo.csv
lookupFile.close()

#########################################################################################################
# H. Record Open Inquires

# Counters 
totalOpen = 0
openInquiresClosed = 0

# dictionary containing key thread id and value subject of all messages in this week's inbox. 
inboxIDs = {}

# Read in messages from the inbox and ad to inboxIDs

for message in mailbox.mbox('Inbox.mbox'):
    if i == 0:
        print "Reading Inbox.mbox"
        i += 1

    threadID = message['X-GM-THRID']
    subject = message['Subject']
    date = message['Date']

    # If any of the field is blank convert from None to "" to avoid errors
    if threadID is None:
        print "Found Thread ID is none"
        threadID = ""
    if subject is None:
        subject = ""
    if date is None:
        date = ""

    inboxIDs[threadID] = subject

# Compare openInquires against inboxIDs to determine if any openInquires have been closed out in the
# past week. 

toDelete = []

# If a thread in openInquires is open and a goodThread increase and in inboxIDs increase totalOpen
# If the thread is in openInquires and a good thread but not in inboxIDs increase openInquiresClosed
# add add to toDelete for deletion from openInquires
# Else if the thread is not in inboxIDs delete the thread. (This gets rid of threads which were closed
# but did not need to be counted.)

for thread in openInquires:
    if openInquires[thread].open and openInquires[thread].good:
        if thread in inboxIDs:
            totalOpen +=1
        else:
            openInquiresClosed += 1
            toDelete.append(thread)
    elif thread not in inboxIDs:
        toDelete.append(thread)

# Delete threads from openInquires
for thread in toDelete:
    del openInquires[thread]

# Write information to open.txt

writeOpenFile = open('Run Files\open.txt', 'wb')

print "Recording open inquiries...\n"
for thread in openInquires:
    writeOpenFile.write(thread + "\n")
    writeOpenFile.write(openInquires[thread].subject + "\n")
    if openInquires[thread].open:
        writeOpenFile.write("Open" + "\n")
    else:
        writeOpenFile.write("closed" + "\n")
    if openInquires[thread].good:
        writeOpenFile.write("Y" + "\n")
    else:
        writeOpenFile.write("N" + "\n")

writeOpenFile.close()

#########################################################################################################
# I. Record Statistics
# Combine Change Requests
# Note. There may be a problem with Change Request - Access Level not exporting. Don't 
# really know whats going on since there isn't any info in mail.csv. For now I will add 
# them bu hand. 
changeRequests = statsLabels["Change Request"].count
changeRequestsAccess = statsLabels["Change Request - Access Level"].count
del statsLabels["Change Request - Access Level"]
statsLabels["Change Request"].count = changeRequests + changeRequestsAccess

# Combine Issue labels
issue = statsLabels["Issue"].count
issuePriority = statsLabels["Issue"].priority
system = statsLabels["System Access Issue"].count
pdf = statsLabels["Issue/PDF"].count
del statsLabels["Issue"]
del statsLabels["System Access Issue"]
del statsLabels["Issue/PDF"]
statsLabels["Issues"] = statInfo(issue + system + pdf, issuePriority)

# Combine CITI Labels
citi1 = statsLabels["CITI Integration"].count
citiPriority = statsLabels["CITI Integration"].priority
citi2 = statsLabels["CITI Interface Errors"].count
del statsLabels["CITI Integration"]
del statsLabels["CITI Interface Errors"]
statsLabels["CITI"] = statInfo(citi1 + citi2, citiPriority)

# Add New Orgs
statsLabels["New Organizations"] = statInfo(newOrgs, priority + 103)

# Add Sales Pings with webform information
salesPings = statsLabels["Sales Pings"].count
statsLabels["Sales Pings"].count = str(salesPings) + "(" + str(webForm) + ")"

# Add Pings
statsLabels["User Inquiries"] = statInfo(pingInquiries, priority + 101)
statsLabels["Demo Requests"] = statInfo(pingDemos, priority + 102)

# Add Voicemails
statsLabels["Voicemails"] = statInfo(voicemails, priority + 104)

# Total Pings, NonPings and Sales Inquires
totalPings = statsLabels["User Inquiries"].count + statsLabels["Demo Requests"].count +statsLabels["Welcome to Support"].count + statsLabels["Welcome Ping"].count + statsLabels["Sales"].count + statsLabels["New Organizations"].count + salesPings +  statsLabels["Voicemails"].count

statsLabels["Total Pings"] = statInfo(totalPings, priority + 105)

totalNonPings = statsLabels["Product Enhancement"].count + statsLabels["WIRB"].count + statsLabels["Funding Status"].count + statsLabels["Alerts"].count + statsLabels["CITI"].count + statsLabels["IRBNet Resources"].count + statsLabels["Reports"].count + statsLabels["Smart Forms"].count + statsLabels["Stamps"].count + statsLabels["Letters"].count + statsLabels["Information"].count + statsLabels["Change Request"].count + statsLabels["Issues"].count 
statsLabels["Total Non-Pings"] = statInfo(totalNonPings, (issuePriority + 0.5))

#Adding Total New Inquiries
statsLabels["Overall Total New Inquiries"] = statInfo(totalPings + totalNonPings, 200)


# Add New Open inquiries
statsLabels["New Open Inquiries"] = statInfo(newOpen, 201)

# Add New Closed inquiries
statsLabels["New Closed Inquiries"] = statInfo(newClosed, 202)

# Add Existing Open inquiries
statsLabels["Existing Open Inquiries Closed"] = statInfo(openInquiresClosed, 203)

# Add Total Open inquiries
statsLabels["Total Open Inquiries"] = statInfo(totalOpen, 204)

# Add Total Closed inquiries
totalClosed = newClosed + openInquiresClosed
statsLabels["Total Closed Inquiries"] = statInfo(totalClosed, 205)

# Add Call Information
enrollmentCallRange = 'Call_Info' # Named Range for Call Information

try:
    callStats = SHEETS_API.spreadsheets().values().get(spreadsheetId=ENROLLMENT_DASHBOARD_ID,
        range=enrollmentCallRange, majorDimension='COLUMNS').execute().get('values',[])
    sessions = callStats[0][0]
    salesCalls = callStats[1][0]
    salesDemos = callStats[2][0]
    if len(callStats) > 3:
        salesDemoInstiutions = str(callStats[3][0])
    else:
        salesDemoInstiutions = ""
except HttpError:
    raw_input("There was an error reading the Enrollment DashBoard. Please let Stephan know. Press enter to continue.")
    sessions = raw_input("\n\nEnter the Total # of Sessions...  ")
    salesCalls = raw_input("Enter the Total # of Sales Calls...  ")
    salesDemos = raw_input("Enter the Total # of Sales Demos...  ")
    salesDemoInstiutions = raw_input("Enter Sales Demo institutions...  ")
except IndexError:
    raw_input("Ben's call information could not be read. Please remind him to enter his calls. Press enter to continue.")
    sessions = raw_input("\n\nEnter the Total # of Sessions...  ")
    salesCalls = raw_input("Enter the Total # of Sales Calls...  ")
    salesDemos = raw_input("Enter the Total # of Sales Demos...  ")
    salesDemoInstiutions = raw_input("Enter Sales Demo institutions...  ")


statsLabels["Total # of Sessions"] = statInfo(sessions, 300)
statsLabels["Total # of Sales Calls"] = statInfo(salesCalls, 301)
statsLabels["Total # of Demo Calls"] = statInfo(salesDemos, 302)


# Add Spacers and other formatting for easy copy and paste
statsLabels[" "] = statInfo("", (issuePriority + 0.6))
statsLabels["Category "] = statInfo("", (issuePriority + 0.7))
statsLabels["     "] = statInfo("",  statsLabels["Total Pings"].priority + 0.1)
statsLabels["      "] = statInfo("",  statsLabels["Total Pings"].priority + 0.2)
statsLabels["       "] = statInfo("",  statsLabels["Total Closed Inquiries"].priority + 0.3)

# Write information to 'stats.csv'
print "Writing Stats info...\n"

# toSort is used since sortStats not compatible with dictionary
toSort = []
for stat in statsLabels:
    toSort.append(stat)

# Creates a list of requests for the sheets API. Each entry in statsColumn in a row in the google sheet
# each value in "values" is a cell in the row. 
statsColumn = [{"values":[{ "userEnteredValue": {"stringValue": str(time.strftime("%m/%d/%Y"))}}]}]

# The call addStatValueToColumn(int i, str valueType) is used to create a Sheets API compatible request 
# for each stat. This method essentially builds a column. 
def addStatValueToColumn(i, valueType):
    return {
            "values":[{
                     "userEnteredValue": {valueType: i}
                     }]
            }

# Add stats to statColumn 
for stat in sortStats(toSort):
    toAdd = statsLabels[stat].count
    if type(toAdd) is int:
        valueType = "numberValue"
    else:
        valueType = "stringValue"
    statsColumn.append(addStatValueToColumn(toAdd, valueType))


# Sheets API request body. Adds a new column to the Weekly Support Stats Sheet and updates it 
# with new stat values. 
batchUpdateBody = {
  "requests": [
    {
      "insertDimension": {
        "inheritFromBefore": False,
        "range": {
          "dimension": "COLUMNS",
          "startIndex": 2,
          "endIndex": 3,
          "sheetId": 0
        }
      }
    }
    ,
    {
      "updateCells": {
        "range": {
          "startRowIndex": 0,
          "endRowIndex": 39,
          "sheetId": 0,
          "startColumnIndex": 2,
          "endColumnIndex": 3
        },
        "rows": statsColumn,
        "fields": "*"
      }
    }
  ]
}

# Executes the stats update request
SHEETS_API.spreadsheets().batchUpdate(spreadsheetId=WEEKLY_STATS_SHEET_ID, body=batchUpdateBody).execute()

######################################################################################################
# J. Member Specific Stats info. 

# Now we'll update  last contact info and cumulative member stats

newMemberStats = []

# Return the A1 notation for a given index. getA1olumnNotation(0) -> A
def getA1ColumnNotation(i):
    ALPHABET = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    j = i-1
    if j < 26:
        return ALPHABET[j]
    else:
        return ALPHABET[j / 26 -1] + ALPHABET[j % 26]


def addMemberStats(mem):
    result = []
    result.append(mem)
    result.append(members[mem].lastContact)
    result.append(members[mem].phone)
    for stat in members[mem].stats:
        result.append(int(stat))
    return result

for mem in sorted(members):
    newMemberStats.append(addMemberStats(mem))

newMemberStatsBody = {
    "values": newMemberStats,
    "majorDimension": "ROWS"
}

memberInfoUpdateRange = MEMBER_STATS_SHEET + '!A2:' + getA1ColumnNotation(NUMBER_OF_MEMBER_STATS + 3)
try:
    SHEETS_API.spreadsheets().values().update(spreadsheetId=SPREADSHEET_ID, body=newMemberStatsBody,
        range=memberInfoUpdateRange, valueInputOption="USER_ENTERED",
        responseValueRenderOption="FORMATTED_VALUE").execute()
except HttpError:
    print "Unable to write Member Specific Stats information"
    sys.exit()

sortRequest = {
    "requests":[ {
      "sortRange": {
        "range": {
          "sheetId": MEMBER_STATS_SHEET_ID,
          "startRowIndex": 1,
          "endRowIndex": 1000,
          "endColumnIndex": 1000,
          "startColumnIndex": 0
        },
        "sortSpecs": [
          {
            "dimensionIndex": 0,
            "sortOrder": "ASCENDING"
          }
        ]
      }
    }]
}
SHEETS_API.spreadsheets().batchUpdate(spreadsheetId=SPREADSHEET_ID, body=sortRequest).execute()

######################################################################################################
# K. Write updated Admin specific data

missedAdminsOut = open("Run Files\Admin Contacts Not Tracked.csv", 'wb')
missedAdminWriter = csv.writer(missedAdminsOut)

for admin in missedAdmins:
    missedAdminWriter.writerow([admin, missedAdmins[admin]])
# This hurts my soul
updatedAdminDates = [""] * len(admins)

for admin in admins:
    updatedAdminDates[admins[admin].row] = [admins[admin].checkIn, admins[admin].lastContact]

newAdminBody = {
    "majorDimension": "ROWS",
    "values": updatedAdminDates
}

updateRange = 'Admin_Contact_Info'
SHEETS_API.spreadsheets().values().update(spreadsheetId=SPREADSHEET_ID, range=updateRange, 
    valueInputOption="RAW", body=newAdminBody).execute()

missedAdminsOut.close()

######################################################################################################
# L. Write the Weekly Stats email 


def createMessage(sender, to, subject, text):
    message = MIMEText(text)
    message['to'] = to
    message['from'] = sender
    message['subject'] = subject
    return {'raw': base64.urlsafe_b64encode(message.as_string())}

def createDraft(service, userID, messageBody):
    try:
        message = {'message': messageBody}
        draft = service.users().drafts().create(userId=userID, body=message).execute()
        return draft
    except HttpError as error:
        print 'An error occurred. Unable to write draft: %s' % error
        return None

def sendDraft(service, userID, draft):
    service.users().drafts().send(userId=userID, body={ 'id':draft['id'] }).execute()


initials = raw_input("Enter your initials...  ")
today1 = time.strftime("%m/%d/%Y")
today = time.strftime("%Y_%m_%d")
todayWeekday = time.strftime("%a")

cutoffDay = time.strptime(cutoff, "%m/%d/%Y").tm_wday
cutoffWeekday = calendar.day_name[cutoffDay]

try:
    retentionCalls = SHEETS_API.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID, range='Weekly_Check_In').execute().get('values',[])
except:
    print "Could not read number of check in calls."
    retentionCalls = raw_input("Enter the number of check in calls this week (cell B2 on Chart Data tab of Retention sheet).   ")


email = "Andy, \n\n"
email = email + "Below are the requested statistics from " + cutoffWeekday + ", " + cutoff + " to " + todayWeekday  + ", " + today1 + ":\n\n"
pingsLessSales = str(totalPings - salesPings)
email = email + "Pings (includes New Orgs; no Sales Pings): " + pingsLessSales + "\n"
email = email + "Non-Pings: " + str(totalNonPings) + "\n"
email = email + "Sales Inquiries: " + statsLabels["Sales Pings"].count + "\n"
total = str(totalPings + totalNonPings)
email = email + "Overall Total New Inquiries: " + total + "\n\n"
email = email + "Total New Inquiries (Non-Pings) Currently Open: " + str(newOpen) + "\n"
email = email + "Total New Inquiries (Non-Pings) Closed: " + str(newClosed) + "\n"
email = email + "Total Existing Open Inquiries Closed: " + str(openInquiresClosed) + "\n"
email = email + "Total Open Inquiries: " + str(totalOpen) + "\n"
email = email + "Total Closed Inquiries: " + str(totalClosed) + "\n\n"
email = email + "Total # of Sessions: " + str(sessions) + "\n"
email = email + "Total # of Sales Calls: " + str(salesCalls) + "\n"
email = email + "Total # of Demo Calls: " + str(salesDemos) + " " + str(salesDemoInstiutions) + "\n"
email = email + "Total # of Retention Calls: " + str(retentionCalls[0][0]) + "\n\n"
email = email + "Cumulative Weekly Statistics can be be accessed via the following sheet:\n"
email = email + "https://docs.google.com/a/irbnet.org/spreadsheets/d/12wQxfv5EOEEsi3zCFwwwAq05SAgvzXoHRZbD33-TQ3o/edit?usp=sharing\n\n"
email = email + "Retention information can be accessed via the following sheet\n"
email = email + "https://docs.google.com/spreadsheets/d/1mkxL43rqDyBZ6T8TIzg1_OQKhVNjefvYTDg9noC18j4/edit#gid=983793166\n\n"
email = email + "Let me know if you have any questions. Thanks!\n\n"
email = email + initials


try:
    message = createMessage("me", STATS_EMAIL, "Stats as of " + today1 + "\n\n", email)
    draft = createDraft(MAIL_API, "me", message)
    sendDraft(MAIL_API, "me", draft)

except:
    emailFile = 'stats_email_' + today + '.txt'

    emailOut = open(emailFile, 'w')

    print "Writing to " + emailFile + "....\n\n"

    emailOut.write("Stats as of " + today1 + "\n\n")
    emailOut.write(email)
    emailOut.close()
    raw_input("Failed to write or send email. A text file has been created with the content of the email. Please send the email to Andy and notify to Stephan of the error. To continue press enter.")

######################################################################################################
# K. Finally

# Duplicate Current Tab in Enrollment Dashboard. 
print "Duplicating Current Tab on Enrollment Dashboard"
try:
    NEW_TITLE = str(today1)[:5]

    batchUpdateRequest = {
        'requests': [
        {'duplicateSheet': {
            'sourceSheetId': CURRENT_SHEET_ID,
            'insertSheetIndex': 2,
            'newSheetName': NEW_TITLE # The new sheet will be renamed to the last time stats was run
            }
        }
        ]
    }

    SHEETS_API.spreadsheets().batchUpdate(spreadsheetId=ENROLLMENT_DASHBOARD_ID, body=batchUpdateRequest).execute()

    #Remove formulas from the duplicate tab
    enrollmentValues = SHEETS_API.spreadsheets().values().get(spreadsheetId=ENROLLMENT_DASHBOARD_ID, range=NEW_TITLE+'!A:A',valueRenderOption='UNFORMATTED_VALUE', majorDimension='COLUMNS').execute().get('values',[])

    updateBody = {
        'values':enrollmentValues,
        'majorDimension':'COLUMNS',
        'range': NEW_TITLE+'!A:A'
    }

    SHEETS_API.spreadsheets().values().update(spreadsheetId=ENROLLMENT_DASHBOARD_ID, valueInputOption='RAW', range=NEW_TITLE+'!A:A', body=updateBody).execute()
    SHEETS_API.spreadsheets().values().clear(spreadsheetId=ENROLLMENT_DASHBOARD_ID, range='Current_Calls', body={}).execute()
    SHEETS_API.spreadsheets().values().clear(spreadsheetId=ENROLLMENT_DASHBOARD_ID, range='Bens_Calls', body={}).execute()

    namedRanges = SHEETS_API.spreadsheets().get(spreadsheetId=ENROLLMENT_DASHBOARD_ID, ranges=NEW_TITLE).execute().get('namedRanges', [])


    ids = []
    namedRangeDeleteRequest = []
    for namedRange in namedRanges:
        id = namedRange.get('namedRangeId', [])
        namedRangeDeleteRequest.append({"deleteNamedRange": {"namedRangeId": id}})


    namedRangeDeleteRequestBody = {"requests": namedRangeDeleteRequest}

    SHEETS_API.spreadsheets().batchUpdate(spreadsheetId=ENROLLMENT_DASHBOARD_ID, body=namedRangeDeleteRequestBody).execute()

except HttpError as e:
    print e
    traceback.print_exception(HttpError, e)
    raw_input("The Current tab in the Enrollment Dashboard could not be duplicated. Please duplicate the tab manually and remove formulas from column A of the duplicated sheet using the paste as values tool. Press enter to continue")

# Remove mbox files
if not KEEP:
    print "\nRemoving mbox files...\n"
    os.remove(MBOX)
    os.remove('Inbox.mbox')

print "End!"