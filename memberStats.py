# name: memberStats.py
# author: Stephan Halarewicz
# date: 02/23/2018
#
# NEW IN THIS VERSION
# 	1. Version Control will now be through git (Repository: "TODO")
#	2. TODO: Integration of Google Sheets. Member retention and stats information is 
# 	   now be stored in 'Retention.gsheet'. Information will be written directly to this sheet
# 	   through use of the Google Sheets API. Data Support files ('mail.csv', 'formatted_mail.csv'
# 	   etc.) will still be stored in '.csv' format. These file will be stored in a directory 'Output'
# 	3. 'Weekly Support Stats.gsheet' will be updated automatically
#	4. TODO: Member stats will include a rolling average of how often they have contacted Support 
# 	   (weekly, monthly, quarterly, annually)
# 	5. TODO: Exceptions thrown if an error encountered to prevent data corruption
#	6. Python 3 will no longer be able to run this script as it is not supported by the Google Sheets API.
#	7. Use an argument parser
# 	8. TODO: Delete mbox files
#	9. TODO: Draft email to Andy
#
# REQUIREMENTS
# 	1. Python 2.7 
#  
#
# EXECUTE WITH: python memberStats filename.mbox -a | -n -i

import csv
import mailbox
import sys
import time
from recordtype import *
import argparse
import os.path

def mboxFile(fname):
	ext = os.path.splitext(fname)[1][1:]
	if not ext == "mbox":
		parser.error("Not a valid .mbox file")
	return fname

# Build args parser to validate mbox file-type and except optional arguments
parser = argparse.ArgumentParser(description='Run weekly stats')

# Required. Check for mbox filetype
parser.add_argument("mbox_file", type=lambda s:mboxFile(s), help="mbox file for which stats will be gathered")
# Optional: -i 1000 will print every 1000 email to help track progress
parser.add_argument("-i", type=int, default=0, help="print every ith email read")

# Optional: -a | -n Automatically count or skip all internal emails. 
group = parser.add_mutually_exclusive_group()
group.add_argument("-a", "--all", action="store_true", help="count all internal threads")
group.add_argument("-n", "--none", action="store_true", help="skip all internal threads")
args = parser.parse_args()

# Store arguments for later use
MBOX = args.mbox_file
COUNT_ALL = args.all
COUNT_EVERY = args.i
COUNT_NONE = args.none

print "PARAMETERS:"
print "    STATS FILE: " + MBOX
if COUNT_ALL:
	print "    COUNT_ALL: All internal threads will be counted"
if COUNT_NONE:
	print "    COUNT_NONE: No internal threads will be counted"
if COUNT_EVERY > 0:
	print "    DISPLAY EVERY: Every " + str(COUNT_EVERY) +" emails"
	

###########################################################################################

# A. Obtain Authentication or credentials to access Google Sheets

