import csv
import time
from tools import config
from tools.lib import mail, members, stats, util, googleAPI
from googleapiclient.errors import HttpError

# Read in existing member stats and stat labels from the Retention sheet
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

# Read in Admin Data from the Retention sheet
try:
    admins, admin_emails = members.Admin.read_admins(config.ADMIN_SHEET, config.SPREADSHEET_ID, googleAPI.SHEETS_API)
except HttpError, e:
    util.print_error("Error: Failed to read Admin data. Please report and resolve. Then re-run stats.")
    raise e

threads = {}

if not config.SKIP:
    # Obtain mail from Support Inbox and begin thread counting
    i = 0

    mail_out, fmail_out, mail_writer, fmail_writer = None, None, None, None
    if config.DEBUG:
        #  Open log files
        mail_out = open('tools/logs/mail.csv', 'wb')
        fmail_out = open('tools/logs/formatted_mail.csv', 'wb')
        mail_writer = csv.writer(mail_out)
        mail_writer.writerow(['Thread ID', 'Date', 'From', 'To', 'Subject', 'X-Gmail-Labels'])
        fmail_writer = csv.writer(fmail_out)
        fmail_writer.writerow(['Thread ID', 'Date', 'From', 'To', 'Subject', 'Labels'])

    print "\nReading Support Inbox..."
    inbox = googleAPI.get_messages(googleAPI.SUPPORT_MAIL_API, "me", "label:Inbox")
    print "...done"

    try:
        print "Reading open inquires"
        open_inquiries = mail.OpenInquiry.from_file("tools/open.txt")
        print "...done"
    except IOError:
        util.print_error('ERROR: open.txt not found or not formatted properly. '
                         'Please check to see if the tools folder contains open.txt')
        print 'This file will need to be reconstructed. You will be asked to go through the current inbox to rebuild ' \
              'the file.'
        print 'Alternatively, you may locate/restore the prior version of open.txt to the tools folder and ' \
              're-run the script. This is recommended.'
        raw_input("Press enter to rebuild the file OR exit the script to locate and restore the file.")
        open_inquiries = mail.OpenInquiry.from_current_inbox(inbox)
        print 'Done reading inbox for open inquiries.'

    print "Reading stats label...."
    gmail_messages = googleAPI.get_messages(googleAPI.SUPPORT_MAIL_API, 'me', 'label:stats')
    print "...done"

    print "Building thread data..."
    for message in gmail_messages:
        msg = mail.Message(message)
        msg_id = msg.get_thread_id()

        if config.DEBUG:
            mail_writer.writerow([msg_id, message['Date'], message['From'], message['To'],
                                  message['Subject'], message['X-Gmail-Labels']])

            fmail_writer.writerow([msg_id, msg.get_date(), msg.get_from_address(), msg.get_to(),
                                   msg.get_subject(), msg.get_labels()])

        if msg.is_spam() or msg.is_idea():
            # Spam and ideas should never count
            continue

        if msg_id not in threads:
            threads[msg_id] = mail.Thread(msg)
        else:
            threads[msg_id].add_message(msg)

        if msg.get_from_address() in admin_emails:
            admin = admins[admin_emails[msg.get_from_address()]]
            admin.update_last_contact(msg.get_date())
            if "check-in call" in msg.get_labels():
                admin.update_check_in(msg.get_date())

        if not config.COUNT_EVERY == 0 and i % config.COUNT_EVERY == 0:
            print i, msg
        i += 1

    if config.DEBUG:
        fmail_out.close()
        mail_out.close()
    print "...done"

    print "Evaluating threads..."
    for thread in threads:
        trd = threads[thread]

        for mem in trd.get_members():  # This isn't 100% accurate but good threads with multiple members are rare.
            member_data[mem].update_last_contact(trd.get_last_contact_date())
            if trd.is_check_in():
                member_data[mem].update_check_in(trd.get_check_in_date())

    print "...done"

    print "Counting stats..."
    new_open_inquires = stats.count_stats(threads, member_data)

    # num_open = Number of threads currently open excluding new open inquires
    # num_closed = Number of threads closed since the last time stats were run.
    num_open, num_closed = mail.OpenInquiry.update(open_inquiries, new_open_inquires, inbox)
    stats.count_open(num_open)
    stats.count_existing_closed(num_closed)
    print "...done"

print "Updating Weekly Support Stats gsheet..."
# Combine and format stats in prep for writing. Combines changes requests, CITI, Issues, Sales Ping. Calculates totals
stats.format_stats()

# Update weekly support stats gsheet
stats.update_weekly_support_stats(googleAPI.SHEETS_API, config.WEEKLY_STATS_SHEET_ID)
print "...done"


# Update member and admin stats
try:
    print "Updating Retention ghseet..."
    new_mem_data = map(members.Member.create_stat_row, member_data.values())

    mem_request = googleAPI.update_request(config.MEMBER_STATS_SHEET_ID, new_mem_data,
                                           1, len(member_data) + 1, 0, len(stat_labels) + 3)
    sort_request = googleAPI.sort_request(config.MEMBER_STATS_SHEET_ID, 0, 1, 1000, 0, 1000)

    new_admin_data = map(lambda adm: members.Admin.create_stat_row(admins[adm]), sorted(admins.keys()))
    admin_request = googleAPI.update_request(config.ADMIN_SHEET_ID, new_admin_data, 1, len(new_admin_data) + 1, 4, 6)

    # If on request fails neither member or admin data wil lbe updated.
    googleAPI.spreadsheet_batch_update(googleAPI.SHEETS_API, config.SPREADSHEET_ID,
                                       [mem_request, sort_request, admin_request])

    print "...done"

except HttpError, e:
    util.print_error("Error: Failed to update Member or Admin. Report and resolve error then re-run stats")
    raise e

# Duplicate current tab on enrolment dash)
new_title = time.strftime('%m/%d')

try:
    print "Updating Enrollment Dashboard..."
    duplicate_request = googleAPI.duplicate_sheet_request(config.CURRENT_SHEET_ID, new_title, 3)
    googleAPI.spreadsheet_batch_update(googleAPI.SHEETS_API, config.ENROLLMENT_DASHBOARD_ID, [duplicate_request])
    print "...done"
except HttpError:
    util.print_error("Error: Failed to duplicate current tab on Enrollment Dashboard. See steps below.")
    print '1. Duplicate Tab and rename as ' + new_title + ". If a conflict tab exists, rename the conflict tab."
    print '2. Copy all cells on duplicated tab and paste as values to remove forumulas'
    print '3. Go to "Data -> Named Ranges" and remove all named ranges associated with the duplicated sheet'
    print '4. Delete all call information on the "Current" Tab of the Enrollment Dashboard'
    raw_input("Press enter to continue.")

else:
    # Try each request individually in case one fails.
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
print "Sending Stats email..."
subject = "Stats as of " + time.strftime("%m/%d/%Y")
email_body = stats.draft_message(config.CUTOFF)

try:
    googleAPI.send_message(googleAPI.MAIL_API, "me", config.STATS_EMAIL, subject, email_body)
except HttpError:
    util.print_error("Error: Failed to send email to Andy. Please use text in email.txt or text printed to terminal")
    raw_input("Press enter to continue.")
print "...done"
