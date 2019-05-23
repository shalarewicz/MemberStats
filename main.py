import sys
import csv
import mailbox
import time

import config
import googleAPI
import members
import stats
import mail
import util

###########################################################################################

# Read in existing member stats and stat labels
member_data, stat_labels = \
    members.Member.read_members(config.MEMBER_STATS_SHEET, config.SPREADSHEET_ID, googleAPI.SHEETS_API)
stats.extract_labels(stat_labels)

print "\nThe following Statistics will be determined..."
for stat in sorted(stats.STAT_LABELS):
    print stat

admins, admin_emails = members.Admin.read_admins(config.ADMIN_SHEET, config.SPREADSHEET_ID, googleAPI.SHEETS_API)

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
    print "\nPreparing " + config.MBOX
    i = 0

    #  Open log files
    mail_out = open("Test\\" + "mail.csv", "wb")
    fmail_out = open("Test\\" + 'formatted_mail.csv', 'wb')
    mail_writer = csv.writer(mail_out)
    mail_writer.writerow(['Thread ID', 'Date', 'From', 'To', 'Subject', 'X-Gmail-Labels'])
    fmail_writer = csv.writer(fmail_out)
    fmail_writer.writerow(['Thread ID', 'Date', 'From', 'To', 'Subject', 'Labels'])

    gmail_messages = googleAPI.get_messages(googleAPI.SUPPORT_MAIL_API, 'me',
                                      'after:2019/05/16 before:2019/05/23 -label:no-reply -in:drafts')
    # for message in gmail_messages:
    for message in mailbox.mbox(config.MBOX):  # TODO Use Mail API to query support directly
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
    mail.OpenInquiry.update(open_inquiries, new_open_inquires, 'Inbox.mbox')

    # Combine and format in prep for writing
    stats.format_stats()

    # Update weekly support stats
    stats.update_weekly_support_stats(googleAPI.SHEETS_API, config.WEEKLY_STATS_SHEET_ID)

    # Update member stats
    update_range = config.MEMBER_STATS_SHEET + '!A2:' + util.get_a1_column_notation(len(stat_labels)+3)
    googleAPI.update_range(googleAPI.SHEETS_API, config.SPREADSHEET_ID, update_range,
                           map(members.Member.create_stat_row, member_data.values()))

    # Update admin dates
    googleAPI.update_range(googleAPI.SHEETS_API, config.SPREADSHEET_ID, 'Admin_Contact_Info',
                           map(lambda adm: members.Admin.create_stat_row(admins[adm]), sorted(admins.keys())), 'RAW')

    # Draft an email
    subject = "Stats as of " + time.strftime("%m/%d/%Y") + "\n\n"
    email_body = stats.draft_message(cutoff)

    googleAPI.send_message(googleAPI.MAIL_API, "me", config.STATS_EMAIL, subject, email_body)

# Open the mbox file and parse each message
    # If this is the first message in the thread create a new thread
    # Else add the message to the thread AND the message is good THEN
