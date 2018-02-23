# name: contactStats.py
# author: Stephan Halarewicz
# date: 08/03/2017
#
# NEW IN THIS VERSION
#	1. Added code for obtaining admin specific data 
#	2. Outputs 'Support Outreach Administrators.csv' and 'Admin Contacts not tracked.csv'
#
# Note that this program runs using Python 2 and not Python 3. I recommend installing
# Python 2.7. 
#
# Execute with python contact_stats.py filename.mbox filename.csv someInt
# 	- flag with -all to automatically count all emails (bypass asking user)
# The Call python contact_stats.py filename.mbox member_info.csv 100 will parse through 
# the mbox file and print every 100th email to help track progress. 
# Result is six files:
#	1. 'mail.csv' - Raw data file with headers Thread ID, Date, From, To, Subject, 
#      X-Gmail-Labels
#	2. 'formatted_mail.csv' - Thread ID, Formatted Date (MM/DD/YYYY), From, To, Subject, 
#	 	Member Labels
#	3. 'stats_YYYY_MM_DD.csv' - Counted statistics. Stats are read in from statsLabels.txt
#   4. 'thread_lookup_info.csv' - Contains email data for counted stats for cross-referencing 
# 		and troubleshooting.
# 	5. 'member_info_YYYY_MM_DD.csv' - Dates of Last Contact for each member as well as cumulative 
# 	    statistics for each member. Note that these include general information for emails before
# 		February 17th as these emails weren't filtered to account for multiple threads and any 
# 		internal emails. 
#	6. 'stats_email_YYYY_MM_DD.txt' - Contains text for the email which should be sent to Andy by 4:30
#	   	pm on Thursdays
# 	7. 'Support Outreach Administrators.csv' - Contains contact info for Support Portal Administrators
# 		as well as the dates which they last contacted support and last had a check in call. 
#	8. 'Admin Contacts not tracked.csv' - List of administrator email address which have emailed into Support
# 		but are not tracked in 'Support Outreach Administrators.csv'
#
# In order to run this program you will need a file with member information formatted with member name
# in column A, last Contact Date in column B, last last check-in date in column C and then any statistic
# labels which should be counted. Note that the file must contain column headers in order to find 
# statistics.  Note that the program will give priority to statistics in leftmost columns has priority. 
# For example, if 'Information' is listed in column C and 'Sales' listed in Column H. Any thread 
# containing both of these labels will be counted as Information. The labels in column headers 
# MUST appear exactly the same as they do in GMail else the label won't be counted. For example, 
# "change Request" or "Change  Request" will not return any matches, but "Change Request" will. 
#
# Here we go. It's a a lot. 

import csv
import mailbox
import sys
import time
from recordtype import *

# Determine if any flags were used during execution. 
	# If an integer, n, is found every nth email will be printed. 
	# If -all is found all emails will automatically be counted. 

flags = {}
flags["intCount"] = 0
flags["allCount"] = False

def isInt(s):
	try:
		int(s)
		return True
	except:
		return False

def findFlags():
	i = 0
	for arg in sys.argv:
		if i == 0 or i == 1 or i == 2 or i == 3:
			i += 1
			continue
		elif isInt(arg):
			i += 1
			flags["intCount"] = int(arg)
		elif arg == '-all':
			i += 1
			flags["allCount"] = True

findFlags()

############################################################################################################

# A. Read in member info and stats to be counted

# 1. open existing member_info.csv file for reading

print "Opening " + sys.argv[2] + " for reading..."

memberReadFile = open(sys.argv[2], "rb")
memberReader = csv.reader(memberReadFile)

print "Done!"

# 2. Read in the Member Name, last contact date, last Check-in date and cumulative stat
#    and add to dictionary 'members' using a record type then print a list of members and stats.

members = {}
memberInfo = recordtype('memberInfo', 'lastContact phone stats')

statsLabels	= {}
statInfo = recordtype('statInfo', 'count priority', default = 0)
memberInfoHeader = []
rownum = 0

for row in memberReader:
	if rownum == 0:
		memberInfoHeader = row
 		colnum = 0
 		priority = 0
		for col in row:
			if colnum == 0 or colnum == 1 or colnum == 2:
				colnum += 1
				continue
			else:
				stat = row[colnum]
				statsLabels[stat] = statInfo(0, priority)
				priority +=1
			colnum += 1
		rownum += 1
		continue
	else:
		colnum = 0
		mem = ""
		date = ""
		phone = ""
		stats = []
		for col in row:
			if colnum ==0:
				mem = row[colnum]
			elif colnum == 1:
				date = row[colnum]
			elif colnum == 2:
				phone = row[colnum]
			else:
				stats.append( int(row[colnum]) )
			colnum += 1
			
		members[mem] = memberInfo(date, phone, stats)

statsLabels["Sales Pings"] = statInfo(0, priority + 30)

print "The current last contact info is...."
for mem in sorted(members):
	print mem + ", " + members[mem].lastContact + ", " + members[mem].phone

print "The following Statistics will be determined..."
for stat in sorted(statsLabels):
	print stat

# 3. close last_contact file
memberReadFile.close()

###########################################################################################################
# B. Read in Administrator specific retention data


# Dictionary for storing administrator data key = name, value = recordtype(adminInfo)
admins = {}
adminEmails = {} #not sure if this is needed key = email value = name used to cross-reference data

# Keep track of emails which aren't tracked. 
missedAdmins = {} 

#Record Types for storing admin info
adminInfo = recordtype('adminInfo', 'org role dept checkIn lastContact email1 email2 email3 phone mobile')

print "Reading Admin Specific Data..."
# Open file for reading
adminIn = open(sys.argv[3], 'rb')
missedAdminsFile = open('Admin Contacts Not Tracked.csv', 'rb')

adminReader = csv.reader(adminIn)
missedAdminReader = csv.reader(missedAdminsFile)

onHeader = True
adminHeader = []
missedAdminHeader = []

# Read in existing admin email addresses which are not tracked
for adminEmail in missedAdminReader:
	if onHeader:
		missedAdminHeader = adminEmail
		onHeader = False
		continue
	missedAdmins[adminEmail[0]] = adminEmail[1]

onHeader = True

# Create an entry in dictionary admins for every admin in sheet. Add each email address to dictionary adminEmails
for admin in adminReader:
	if onHeader:
		adminHeader = admin
		onHeader = False
		continue
	org = admin[0]
	name = admin[1]
	role = admin[2]
	dept = admin[3]
	checkInDate = admin[4]
	lastContactDate = admin[5]
	email1 = admin[6].lower()
	email2 = admin[7].lower()
	email3 = admin[8].lower()
	phone = admin[9]
	mobile = admin[10]

	admins[name] = adminInfo(org, role, dept, checkInDate, lastContactDate, email1, email2, email3, phone, mobile)
	adminEmails[email1] = name
	if not email2 == "":
		adminEmails[email2] = name
		print "Found blank email 2"
	if not email3 == "":
		adminEmails[email3] = name
		print "Found blank email 3"

# Close the file for reading
adminIn.close()
missedAdminsFile.close()
print "Done"
###########################################################################################################
# B. Read in Open Inquires

# Read in current open inquiries from open.txt. 
#
# 	open.txt should have format 
# 	threadID
# 	subject
# 	"Open" or Closed"
#   "Y" or "N" - indicates if thread should be counted. 
# 
# A text file is used to store this information rather than a csv due to Excel truncating thread ID since
# the number is to big. 

# openInquiries used to store information
openInquires = {}
openInfo = recordtype('openInfo', 'subject, open, good')

lastWeekOpen = 0

print "\n\n\nBegin\n"

# Open open.txt for reading
print "Opening open.txt for reading..."
openInquiriesFile = open('open.txt', 'rb')
print "Done!"

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

print "Done!"

openInquiriesFile.close()



###########################################################################################################

# C. Helper Functions

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
	if "owen@irbnet.org" in frm:
		return True
	if "stephan@irbnet.org" in frm:
		return True
	if "andy@irbnet.org" in frm:
		return True
	if "ben@irbnet.org" in frm:
		return True
	if "april@irbnet.org" in frm:
		return True
	if "chris@irbnet.org" in frm:
		return True
	if "jake@irbnet.org" in frm:
		return True
	if "jkaten@irbnet.org" in frm:
		return True
	if "pmcclammer@irbnet.org" in frm:
		return True
	if "kayla@irbnet.org" in frm:
		return True
	if "emily@irbnet.org" in frm:
		return True
	if "deena@irbnet.org" in frm:
		return True
	if "jmestre@irbnet.org" in frm:
		return True


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
	if not flags["allCount"]:
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
	else:
		pass

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

# C. Now for the fun part

# 1. Open mail.csv and formatted_mail.csv for writing then write the headers



print "Opening outfile...."
firstOutFile = open('mail.csv', 'wb')

print "Opened mail.csv. Preparing to write..."
writer = csv.writer(firstOutFile)

print "Ready to write"

print "Opened formatted_mail.csv. Preparing to write..."
outfile = open('formatted_mail.csv', "wb")
datedWriter = csv.writer(outfile)

print "Ready to write"

print "Writing Headers"
writer.writerow(['Thread ID', 'Date', 'From', 'To', 'Subject', 'X-Gmail-Labels'])
datedWriter.writerow(['Thread ID', 'Date', 'From', 'To', 'Subject', 'Stats', 'Members'])
print "Done!"

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
cutoff = raw_input("\nOn which date were stats last run?\ni.e. What is the earliest date for which stats should count\n(MM/DD/YYYY)    ")

# 2. Open mbox file and write raw data to mail.csv. Use Helper functions above to 
# 	 format date and check for any member and stats labels then add appropriate information to
# 	 dictionary "threads"

print "Preparing " + sys.argv[1]
for message in mailbox.mbox(sys.argv[1]):
	if i == 0:
		print "Starting to read messages"
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
			if "norepy@irbnet.org" not in fromAddress:
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
	if not flags["intCount"] == 0 and i % flags["intCount"] == 0:
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
# D. Update date of last contact and check-in date for each member

# 1. compare the most recent date for each thread against the date of last contact
#    for each member and sway the them in necessary.

print "Updating dates of last contact and check-in call dates..."

for thread in threads:
	if threads[thread].memberLabels == []:
		pass
	else:
		for mem in threads[thread].memberLabels:
			if compareDates(threads[thread].date, members[mem].lastContact):
				members[mem].lastContact = threads[thread].date
			if compareDates(threads[thread].checkInDate, members[mem].phone):
				members[mem].phone = threads[thread].checkInDate
print "Done!"

#######################################################################################
# E. Count Statistics

# Open 'threadLookupInfo.csv' for writing
# This is done to cross reference which stat was counted for a thread if there is ever a dispute

print "Opening thread_lookup_info.csv...."

lookupFile = open('threadLookupInfo.csv', 'wb')
lookupWriter = csv.writer(lookupFile)

print "Done!"

# Write headers. 
print "Writing Headers..."

lookupWriter.writerow(["Date", "Subject", "From", "To", "Labels", "Counted Stat", "Closed?"])

print "Done!"

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
# Close threadLookupInfo.csv
print "Closing the lookup file..."

lookupFile.close()

print "Done!"

#########################################################################################################
# F. Record Open Inquires

# Counters 
totalOpen = 0
openInquiresClosed = 0

# dictionary containing key thread id and value subject of all messages in this week's inbox. 
inboxIDs = {}

# Read in messages from the inbox and ad to inboxIDs
print "Opening Inbox.mbox..."

for message in mailbox.mbox('Inbox.mbox'):
	if i == 0:
		print "Starting to read messages"
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
print "Done!"

# Compare openInquires against inboxIDs to determine if any openInquires have been closed out in the
# past week. 

toDelete = []

# If a thread in openInquires is open and a goodThread increase and in inboxIDs increase totalOpen
# If the thread is in openInquires and a good thread but not in inboxIDs increase openInquiresClosed
# add add to toDelete for deletion from openInquires
# Else if the thread is not in inboxIDs delete the thread. (This gets rid of threads which were closed
# but did not need to be counted.)

print "Comparing Open Inquires..."
for thread in openInquires:
	if openInquires[thread].open and openInquires[thread].good:
		if thread in inboxIDs:
			totalOpen +=1
		else:
			openInquiresClosed += 1
			toDelete.append(thread)
	elif thread not in inboxIDs:
		toDelete.append(thread)
print "Done!"

# Delete threads from openInquires
print "Deleting Closed Inquires..."
for thread in toDelete:
	del openInquires[thread]
print "Done!"

# Write information to open.txt
print "Opening open.txt for writing...."

writeOpenFile = open('open.txt', 'wb')

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

print "Done!"

#########################################################################################################
# G. Record Statistics

# Combine Change Requests
# Note. There may be a problem with Change Request - Access Level not exporting. Don't 
# really know whats going on since there isn't any info in mail.csv. For now I will add 
# them bu hand. 
print "Combining Change Requests..."
changeRequests = statsLabels["Change Request"].count
changeRequestsAccess = statsLabels["Change Request - Access Level"].count
del statsLabels["Change Request - Access Level"]
statsLabels["Change Request"].count = changeRequests + changeRequestsAccess
print "Done!"

# Combine Issue labels
print "Combining Issues..."
issue = statsLabels["Issue"].count
issuePriority = statsLabels["Issue"].priority
system = statsLabels["System Access Issue"].count
pdf = statsLabels["Issue/PDF"].count
del statsLabels["Issue"]
del statsLabels["System Access Issue"]
del statsLabels["Issue/PDF"]
statsLabels["Issues"] = statInfo(issue + system + pdf, issuePriority)
print "Done!"

# Combine CITI Labels
print "Combining CITI..."
citi1 = statsLabels["CITI Integration"].count
citiPriority = statsLabels["CITI Integration"].priority
citi2 = statsLabels["CITI Interface Errors"].count
del statsLabels["CITI Integration"]
del statsLabels["CITI Interface Errors"]
statsLabels["CITI"] = statInfo(citi1 + citi2, citiPriority)
print "Done!"

# Add New Orgs
print "Adding New Orgs.."
statsLabels["New Organizations"] = statInfo(newOrgs, priority + 103)
print "Done!"

# Add Sales Pings
print "Combining Sales Pings with webform..."
salesPings = statsLabels["Sales Pings"].count
statsLabels["Sales Pings"].count = str(salesPings) + "(" + str(webForm) + ")"
print "Done!"

# Add Pings
print "Adding Pings..."
statsLabels["User Inquiries"] = statInfo(pingInquiries, priority + 101)
statsLabels["Demo Requests"] = statInfo(pingDemos, priority + 102)
print "Done!"

# Add Voicemails
print "Adding Voicemails..."
statsLabels["Voicemails"] = statInfo(voicemails, priority + 104)
print "Done!"

# Total Pings, NonPings and Sales Inquires
print "Totaling Pings..."
totalPings = statsLabels["User Inquiries"].count + statsLabels["Demo Requests"].count +statsLabels["Welcome to Support"].count + statsLabels["Welcome Ping"].count + statsLabels["Sales"].count + statsLabels["New Organizations"].count + salesPings +  statsLabels["Voicemails"].count
print "Done! Total Pings = " + str(totalPings)

statsLabels["Total Pings"] = statInfo(totalPings, priority + 105)

print "Totaling Non-Pings..."
totalNonPings = statsLabels["Product Enhancement"].count + statsLabels["WIRB"].count + statsLabels["Funding Status"].count + statsLabels["Alerts"].count + statsLabels["CITI"].count + statsLabels["IRBNet Resources"].count + statsLabels["Reports"].count + statsLabels["Smart Forms"].count + statsLabels["Stamps"].count + statsLabels["Letters"].count + statsLabels["Information"].count + statsLabels["Change Request"].count + statsLabels["Issues"].count 
print "Done! Total Non-Pings = " + str(totalNonPings)
statsLabels["Total Non-Pings"] = statInfo(totalNonPings, (issuePriority + 0.5))

#Adding Total New Inquiries
print "Adding Total New Inquiries"
statsLabels["Overall Total New Inquiries"] = statInfo(totalPings + totalNonPings, 200)
print "Done! New Open Inquiries = " + str(newOpen)

# Add New Open inquiries
print "Adding new open inquiries"
statsLabels["New Open Inquiries"] = statInfo(newOpen, 201)
print "Done! New Open Inquiries = " + str(newOpen)

# Add New Closed inquiries
print "Adding new closed inquiries"
statsLabels["New Closed Inquiries"] = statInfo(newClosed, 202)
print "Done! New Closed Inquiries = " + str(newClosed)

# Add Existing Open inquiries
print "Adding existing open inquiries closed"
statsLabels["Existing Open Inquiries Closed"] = statInfo(openInquiresClosed, 203)
print "Done! Existing Open Inquiries = " + str(openInquiresClosed)

# Add Total Open inquiries
print "Adding total open inquiries"
statsLabels["Total Open Inquiries"] = statInfo(totalOpen, 204)
print "Done! Total Open Inquiries = " + str(totalOpen)

# Add Total Closed inquiries
totalClosed = newClosed + openInquiresClosed
print "Adding total closed inquiries"
statsLabels["Total Closed Inquiries"] = statInfo(totalClosed, 205)
print "Done! Total Closed Inquires = " + str(totalClosed)

# Add Call Information
sessions = raw_input("\n\nEnter the Total # of Sessions...  ")
salesCalls = raw_input("Enter the Total # of Sales Calls...  ")
salesDemos = raw_input("Enter the Total # of Sales Demos...  ")

statsLabels["Total # of Sessions"] = statInfo(sessions, 300)
statsLabels["Total # of Sales Calls"] = statInfo(salesCalls, 301)
statsLabels["Total # of Demo Calls"] = statInfo(salesDemos, 302)

# Open new Stats file
today1 = time.strftime("%m/%d/%Y")
today = time.strftime("%Y_%m_%d")
newFile = "stats_" + today + '.csv'

# Add Spacers and other formatting for easy copy and paste
statsLabels["Category"] = statInfo(today1, -1)
statsLabels[" "] = statInfo("", (issuePriority + 0.6))
statsLabels["Category "] = statInfo("", (issuePriority + 0.7))
statsLabels["     "] = statInfo("",  statsLabels["Total Pings"].priority + 0.1)
statsLabels["      "] = statInfo("",  statsLabels["Total Pings"].priority + 0.2)
statsLabels["       "] = statInfo("",  statsLabels["Total Closed Inquiries"].priority + 0.3)

print "\n\nOpening " + newFile + " for writing"

statsFile = open(newFile, "wb")
statsWriter = csv.writer(statsFile)

print "Done"

# Write information to 'stats.csv'
print "Writing Stats info..."

# toSort is used since sortStats not compatible with dictionary
toSort = []
for stat in statsLabels:
	toSort.append(stat)

# Write stats to stats.csv by priority lo --> hi
for stat in sortStats(toSort):
	statsWriter.writerow([stat, statsLabels[stat].count])

print "Done!"


print "Closing stats csv..."
# close stats.csv
statsFile.close()

print "Done!"

######################################################################################################

# F. Member Specific Stats info. 

# Now we'll write updated last contact info as well as updated cumulative member stats and 
# and write a new version of member_info_blank.csv to account for any new institutions

# Name the output file with the current date.
newMemberFile = 'member_info_' + today + '.csv'
newMemberBlank = 'member_info_blank.csv'

print "Opening " + newMemberFile + " and " + newMemberBlank+ " for writing"

out = open(newMemberFile, 'wb')
outWriter = csv.writer(out)

outBlank = open(newMemberBlank, 'wb')
outWriterBlank = csv.writer(outBlank)

print "Done!"

# Write Headers
print "Writing headers...."

outWriter.writerow(memberInfoHeader)
outWriterBlank.writerow(memberInfoHeader)

print "Done!"

# Write Updated Information. 
print "Writing new member information..."
for mem in sorted(members):
	toWrite = [mem, members[mem].lastContact, members[mem].phone]
	outWriterBlank.writerow([mem])
	for stat in members[mem].stats:
		toWrite.append(stat)
	outWriter.writerow(toWrite)

print "Done!"	

print "Closing " + newMemberFile + " and " + newMemberBlank

out.close()
outBlank.close()

print "Done!"	

#######################################################################################
# H. Write updated Admin specific data

newAdminFile = 'Support Outreach Administrators_' + today + '.csv'

print "Opening new Support Outreach File for writing...."
adminOut = open(newAdminFile, 'wb')
missedAdminsOut = open("Admin Contacts Not Tracked.csv", 'wb')

adminWriter = csv.writer(adminOut)
missedAdminWriter = csv.writer(missedAdminsOut)

print "Done! Writing...."

adminWriter.writerow(adminHeader)

for admin in sortedAdmins(admins):
	org = admins[admin].org
	role = admins[admin].role
	dept = admins[admin].dept
	checkInDate = admins[admin].checkIn
	lastContactDate = admins[admin].lastContact
	email1 = admins[admin].email1
	email2 = admins[admin].email2
	email3 = admins[admin].email3
	phone = admins[admin].phone
	mobile = admins[admin].mobile
	adminWriter.writerow([org, admin, role, dept, checkInDate, lastContactDate, email1, email2, email3, phone, mobile])
# Sort missed Admins by if it has been checked. 

# Write header for missed admin file
missedAdminWriter.writerow(missedAdminHeader)

for admin in missedAdmins:
	missedAdminWriter.writerow([admin, missedAdmins[admin]])

print "Done! Closing...!"
adminOut.close()
missedAdminsOut.close()

######################################################################################################

# H. Write the Weekly Stats email 

# The following generates a text file which can be pasted into an email for Andy. 

# open text file for writing
emailFile = 'stats_email_' + today + '.txt'

print "Opening " + emailFile + " for writing..."
emailOut = open(emailFile, 'w')
print "Done"


print "Writing email....\n\n"
initials = raw_input("Enter your initials...  ")

emailOut.write("Stats as of " + today1 + "\n\n")

emailOut.write("Andy, \n\n")
emailOut.write(("Below are the requested statistics from Thursday, " + cutoff + " to Thursday, " + today1 + ":\n\n"))
pingsLessSales = str(totalPings - salesPings)
emailOut.write("Pings (includes New Organizations; does not include Sales Pings): " + pingsLessSales + "\n")
emailOut.write("Non-Pings: " + str(totalNonPings) + "\n")
emailOut.write("Sales Inquiries: " + statsLabels["Sales Pings"].count + "\n")

total = str(totalPings + totalNonPings)
emailOut.write("Overall Total New Inquiries: " + total + "\n\n")

emailOut.write("Total New Inquiries (Non-Pings) Currently Open: " + str(newOpen) + "\n")
emailOut.write("Total New Inquiries (Non-Pings) Closed: " + str(newClosed) + "\n")
emailOut.write("Total Existing Open Inquiries Closed: " + str(openInquiresClosed) + "\n")
emailOut.write("Total Open Inquiries: " + str(totalOpen) + "\n")
emailOut.write("Total Closed Inquiries: " + str(totalClosed) + "\n\n")

emailOut.write("Total # of Sessions: " + str(sessions) + "\n")
emailOut.write("Total # of Sales Calls: " + str(salesCalls) + "\n")
emailOut.write("Total # of Demo Calls: " + str(salesDemos) + "\n\n")

emailOut.write("Cumulative Weekly Statistics can be be accessed via the following sheet:\n")
emailOut.write("https://docs.google.com/a/irbnet.org/spreadsheets/d/12wQxfv5EOEEsi3zCFwwwAq05SAgvzXoHRZbD33-TQ3o/edit?usp=sharing\n\n")

emailOut.write("Let me know if you have any questions. Thanks!\n\n")

emailOut.write(initials)

emailOut.close()


# Finally
print "\n\nEnd!"