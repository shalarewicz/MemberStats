import sys
import csv
import time

import config
import googleAPI
import members
import stats
import mail
import util
from googleapiclient.errors import HttpError

###########################################################################################

# Read in existing member stats and stat labels
member_data = None
stat_labels = None
admins = None
admin_emails = None
try:
    member_data, stat_labels = \
        members.Member.read_members(config.MEMBER_STATS_SHEET, config.SPREADSHEET_ID, googleAPI.SHEETS_API)
    stats.extract_labels(stat_labels)
except HttpError, e:
    util.print_error("Error: Failed to read Member data. Please report and resolve. Then re-run stats.")
    raise e

print "\nThe following Statistics will be determined..."
for stat in sorted(stats.STAT_LABELS):
    print stat

try:
    admins, admin_emails = members.Admin.read_admins(config.ADMIN_SHEET, config.SPREADSHEET_ID, googleAPI.SHEETS_API)
except HttpError, e:
    util.print_error("Error: Failed to read Admin data. Please report and resolve. Then re-run stats.")
    raise e

try:
    open_inquiries = mail.OpenInquiry.from_file("Test/open.txt")  # TODO Update location
except IOError:
    # TODO If open.txt not found reconstruct here using openInbox.py
    print "ERROR: Could not read open.txt. Please ensure the file is formatted properly.\n", \
            "THREAD_ID\n," "SUBJECT"
    sys.exit()

cutoff = util.get_cutoff_date()

messages = {}
threads = {}

if not config.SKIP:
    print "\nReading Support Inbox..."
    i = 0

    #  Open log files
    mail_out = open("Test\\" + "mail.csv", "wb")
    fmail_out = open("Test\\" + 'formatted_mail.csv', 'wb')
    mail_writer = csv.writer(mail_out)
    mail_writer.writerow(['Thread ID', 'Date', 'From', 'To', 'Subject', 'X-Gmail-Labels'])
    fmail_writer = csv.writer(fmail_out)
    fmail_writer.writerow(['Thread ID', 'Date', 'From', 'To', 'Subject', 'Labels'])

    gmail_messages = googleAPI.get_messages(googleAPI.SUPPORT_MAIL_API, 'me', 'label:stats')
    for message in gmail_messages:
        msg = mail.Message(message)
        msg_id = msg.get_thread_id()

        mail_writer.writerow([msg_id, message['Date'], message['From'], message['To'],
                              message['Subject'], message['X-Gmail-Labels']])

        fmail_writer.writerow([msg_id, msg.get_date(), msg.get_from_address(), msg.get_to(),
                              msg.get_subject(), msg.get_labels()])

        if msg.is_spam() or msg.is_idea():
            continue

        if msg_id not in threads:
            threads[msg_id] = mail.Thread(msg)
        else:
            threads[msg_id].add_message(msg)

        if "irbnet.org" not in msg.get_from_address() and msg.get_from_address() in admin_emails:
            a = admins[admin_emails[msg.get_from_address()]]
            a.update_last_contact(msg.get_date())
            if "check-in call" in msg.get_labels():
                a.update_check_in(msg.get_date())

        if not config.COUNT_EVERY == 0 and i % config.COUNT_EVERY == 0:
            print i, msg
        i += 1

    fmail_out.close()
    mail_out.close()

    for thread in threads:
        trd = threads[thread]
        if trd.get_oldest_date() < cutoff:
            # TODO Move to Thread class. if the oldest date is ever before the cutoff then we don't need to count it
            # If the thread has an oldest date before the cutoff change goodThread to false. This can
            # also be configured to ask the user if the thread should be counted. However, there are
            # a surprising amount of threads where this occurs 99% of which should not count so I
            # decided to automatically not count the thread. Therefore, if an admin replies to a thread
            # older than the cutoff with a completely new inquiry the thread won't be counted.
            # I'm OK with this as it does not happen too often.
            # TODO Figure out if the extract pulls all emails in the thread or just those within the
            #  specified timeframe
            trd.dont_count()

        if len(trd.get_members()) > 0:
            for mem in trd.get_members():
                member_data[mem].update_last_contact(trd.get_last_contact_date())
                if trd.is_check_in():
                    member_data[mem].update_check_in(trd.get_check_in_date())

    new_open_inquires = stats.count_stats(threads, member_data)
    inbox = googleAPI.get_messages(googleAPI.SUPPORT_MAIL_API, "me", "label:Inbox")
    mail.OpenInquiry.update(open_inquiries, new_open_inquires, inbox)

# Combine and format in prep for writing
stats.format_stats()

# Update weekly support stats
stats.update_weekly_support_stats(googleAPI.SHEETS_API, config.WEEKLY_STATS_SHEET_ID)

# Update member stats and sort the sheet
try:
    update_range = config.MEMBER_STATS_SHEET + '!A2:' + util.get_a1_column_notation(len(stat_labels)+3)
    googleAPI.update_range(googleAPI.SHEETS_API, config.SPREADSHEET_ID, update_range,
                           map(members.Member.create_stat_row, member_data.values()))
except HttpError, e:
    util.print_error("Error: Failed to update Member Stats. Report and resolve error then re-run stats")
    raise e

try:
    sort_request = googleAPI.sort_request(config.MEMBER_STATS_SHEET_ID, 0, 1, 1000, 0, 1000)
    googleAPI.spreadsheet_batch_update(googleAPI.SHEETS_API, config.SPREADSHEET_ID, [sort_request])
except HttpError:
    util.print_error("Error: Failed to sort 'Member Stats' sheet.")

# Update admin dates
try:
    googleAPI.update_range(googleAPI.SHEETS_API, config.SPREADSHEET_ID, 'Admin_Contact_Info',
                           map(lambda adm: members.Admin.create_stat_row(admins[adm]), sorted(admins.keys())), 'RAW')
except HttpError, e:
    util.print_error("Error: Failed to update admin data. Report and resolve error then re-run stats")
    raise e

# Duplicate current tab on enrolment dash)
new_title = time.strftime('%m/%d')

try:
    duplicate_request = googleAPI.duplicate_sheet_request(config.CURRENT_SHEET_ID, new_title, 2)
    googleAPI.spreadsheet_batch_update(googleAPI.SHEETS_API, config.ENROLLMENT_DASHBOARD_ID, [duplicate_request])
except HttpError:
    util.print_error("Error: Failed to duplicate current tab on Enrollment Dashboard. See steps below.")
    print '1. Duplicate Tab and rename as ' + new_title
    print '2. Copy all cells on duplicated tab and paste as values to remove forumulas'
    print '3. Go to "Data -> Named Ranges" and remove all named ranges associated with the duplicated sheet'
    print '4. Delete all call information on the "Current" Tab of the Enrollment Dashboard'
    raw_input("Press enter to continue.")

else:
    try:
        delete_named_ranges_request = googleAPI.delete_named_range_request(
            googleAPI.SHEETS_API, config.ENROLLMENT_DASHBOARD_ID, new_title)
        googleAPI.spreadsheet_batch_update(googleAPI.SHEETS_API, config.ENROLLMENT_DASHBOARD_ID,
                                           [delete_named_ranges_request])
    except HttpError:
        util.print_error("Error: Failed to delete named ranges on duplicated current tab. See steps below.")
        print 'On the Enrollment Dashboard go to "Data -> Named Ranges" and remove all named ranges associated ' \
              'with the tab ' + new_title
        raw_input("Press enter to continue.")
    try:
        googleAPI.remove_formulas(googleAPI.SHEETS_API, config.ENROLLMENT_DASHBOARD_ID, new_title + '!A:A')
    except HttpError:
        util.print_error("Error: Failed to remove values on duplicated current tab. See steps below.")
        print 'Copy all cells on ' + new_title + ' of the Enrollment Dashboard and paste as values to remove formulas'
        raw_input("Press enter to continue.")
    try:
        googleAPI.clear_ranges(googleAPI.SHEETS_API, config.ENROLLMENT_DASHBOARD_ID, ['Current_Calls', 'Bens_Calls'])
    except HttpError:
        util.print_error("Error: Failed to clear cells on Current tab of Enrollment Dashboard. See steps below.")
        print "Delete all call information on the 'Current' Tab of the Enrollment Dashboard"
        raw_input("Press enter to continue.")

# Draft an email
subject = "Stats as of " + time.strftime("%m/%d/%Y")
email_body = stats.draft_message(cutoff)

try:
    googleAPI.send_message(googleAPI.MAIL_API, "me", config.STATS_EMAIL, subject, email_body)
except HttpError:
    util.print_error("Error: Failed to send email to Andy. Please use text in email.txt or text printed to the terminal")
    raw_input("Press enter to continue.")

# Open the mbox file and parse each message
    # If this is the first message in the thread create a new thread
    # Else add the message to the thread AND the message is good THEN
