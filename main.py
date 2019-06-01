import csv
import time

import config
import googleAPI
import members
import stats
import mail
import util
from googleapiclient.errors import HttpError

# Read in existing member stats and stat labels
member_data, stat_labels, admins, admin_emails = None, None, None, None
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

cutoff = util.get_cutoff_date()

messages = {}
threads = {}

if not config.SKIP:
    print "\nReading Support Inbox..."
    i = 0

    mail_out, fmail_out, mail_writer, fmail_writer = None, None, None, None
    if config.DEBUG:
        #  Open log files
        mail_out = open('config/mail.csv', 'wb')
        fmail_out = open('config/formatted_mail.csv', 'wb')
        mail_writer = csv.writer(mail_out)
        mail_writer.writerow(['Thread ID', 'Date', 'From', 'To', 'Subject', 'X-Gmail-Labels'])
        fmail_writer = csv.writer(fmail_out)
        fmail_writer.writerow(['Thread ID', 'Date', 'From', 'To', 'Subject', 'Labels'])

    inbox = googleAPI.get_messages(googleAPI.SUPPORT_MAIL_API, "me", "label:Inbox")

    try:
        open_inquiries = mail.OpenInquiry.from_file("config/open.txt")
    except IOError:
        util.print_error('ERROR: config/open.txt not found or not formatted properly. '
                         'Please check to see if the config folder contains open.txt')
        print 'This file will need to be reconstructed. You will be asked to go through the current inbox to rebuild ' \
              'the file.'
        print 'Alternatively, you may locate/restore the prior version of open.txt to the config folder and ' \
              're-run the script. This is recommended.'
        raw_input("Press enter to rebuild the file OR exit the script to locate and restore the file.")
        open_inquiries = mail.OpenInquiry.from_current_inbox(inbox)
        print 'Done reading inbox for open inquiries.'

    gmail_messages = googleAPI.get_messages(googleAPI.SUPPORT_MAIL_API, 'me', 'label:stats')

    for message in gmail_messages:
        msg = mail.Message(message)
        msg_id = msg.get_thread_id()

        if config.DEBUG:
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

    if config.DEBUG:
        fmail_out.close()
        mail_out.close()

    for thread in threads:
        trd = threads[thread]
        if trd.get_oldest_date() < cutoff:
            trd.dont_count()

        if len(trd.get_members()) > 0:
            for mem in trd.get_members():
                member_data[mem].update_last_contact(trd.get_last_contact_date())
                if trd.is_check_in():
                    member_data[mem].update_check_in(trd.get_check_in_date())

    new_open_inquires = stats.count_stats(threads, member_data)
    num_open, num_closed = mail.OpenInquiry.update(open_inquiries, new_open_inquires, inbox)
    stats.count_open(num_open)
    stats.count_existing_closed(num_closed)

# Combine and format stats in prep for writing
stats.format_stats()

# Update weekly support stats
stats.update_weekly_support_stats(googleAPI.SHEETS_API, config.WEEKLY_STATS_SHEET_ID)

# Update member and admin stats
try:
    new_mem_data = map(members.Member.create_stat_row, member_data.values())

    mem_request = googleAPI.update_request(config.MEMBER_STATS_SHEET_ID, new_mem_data,
                                           1, len(member_data) + 1, 0, len(stat_labels) + 3)
    sort_request = googleAPI.sort_request(config.MEMBER_STATS_SHEET_ID, 0, 1, 1000, 0, 1000)

    new_admin_data = map(lambda adm: members.Admin.create_stat_row(admins[adm]), sorted(admins.keys()))
    admin_request = googleAPI.update_request(config.ADMIN_SHEET_ID, new_admin_data, 1, len(new_admin_data) + 1, 4, 6)

    googleAPI.spreadsheet_batch_update(googleAPI.SHEETS_API, config.SPREADSHEET_ID,
                                       [mem_request, sort_request, admin_request])

except HttpError, e:
    util.print_error("Error: Failed to update Member Stats. Report and resolve error then re-run stats")
    raise e

# Duplicate current tab on enrolment dash)
new_title = time.strftime('%m/%d')

try:
    duplicate_request = googleAPI.duplicate_sheet_request(config.CURRENT_SHEET_ID, new_title, 2)
    googleAPI.spreadsheet_batch_update(googleAPI.SHEETS_API, config.ENROLLMENT_DASHBOARD_ID, [duplicate_request])
except HttpError:
    util.print_error("Error: Failed to duplicate current tab on Enrollment Dashboard. See steps below.")
    print '1. Duplicate Tab and rename as ' + new_title + ". If a conflict tab exists, rename the conflict tab."
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
    util.print_error("Error: Failed to send email to Andy. Please use text in email.txt or text printed to terminal")
    raw_input("Press enter to continue.")
