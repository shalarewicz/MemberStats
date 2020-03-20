import csv
from tools import config
from tools.lib import mail, members, stats, util, googleAPI
from googleapiclient.errors import HttpError


def read_members(mem_stats_sheet, retention_sheet_id, sheets_api, stat_header_index, short_name_range):
    """
    Read in existing member stats and stat labels from the Retention sheet
    :param mem_stats_sheet: Sheet name for Member Stats tab on sheet specified by retention_sheet_id
    :param retention_sheet_id: Google Sheet ID for Retention sheet. Must contain mem_stats_sheet
    :param sheets_api: Sheets api used to execute the query
    :param stat_header_index: Start index (inclusive) of values that will be returned from the mem_stats_sheet header
    :param short_name_range: Named Range in Retention sheet for short names. If mem_stats_sheet does not contain a name
                            in this range it will be added to the mem_stats_sheet
    :return: member dictionary with (k,v) = (name, Member() object), mem_stats_sheet header[stat_header_index:]
    """
    try:
        member_data, stat_labels = \
            members.Member.read_members(mem_stats_sheet, retention_sheet_id, sheets_api, stat_header_index,
                                        short_name_range)
    except HttpError, e:
        util.print_error("Error: Failed to read Member data. Please report and resolve. Then re-run stats.")
        raise e

    return member_data, stat_labels


def read_admins(admin_sheet, retention_sheet_id, sheets_api):
    # Read in Admin Data from the Retention sheet
    try:
        admins, admin_emails = members.Admin.read_admins(admin_sheet, retention_sheet_id, sheets_api)
    except HttpError, e:
        util.print_error("Error: Failed to read Admin data. Please report and resolve. Then re-run stats.")
        raise e
    return admins, admin_emails


def update_admin_dates(threads, admins, admin_emails):
    for thread in threads:
        for msg in threads[thread].messages:
            if msg.get_from_address() in admin_emails:
                admin = admins[admin_emails[msg.get_from_address()]]
                admin.update_last_contact(msg.get_date())
                if "check-in call" in msg.get_labels():
                    admin.update_check_in(msg.get_date())


def read_inbox(stat_counter, file_base, support_mail_api):
    print "\nReading Support Inbox..."
    inbox = googleAPI.get_messages_from_threads(support_mail_api, "me", "label:Inbox")
    print "...done"

    try:
        print "Reading open inquires"
        open_inquiries = mail.OpenInquiry.from_file('tools/' + file_base + 'open.txt')
        print "...done"
    except IOError:
        util.print_error('ERROR: ' + file_base + 'open.txt not found or not formatted properly. '
                         'Please check to see if the tools folder contains ' + file_base + 'open.txt')
        print 'This file will need to be reconstructed. You will be asked to go through the current inbox to rebuild ' \
              'the file.'
        print 'Alternatively, you may locate/restore the prior version of open.txt to the tools folder and ' \
              're-run the script. This is recommended.'
        raw_input("Press enter to rebuild the file OR exit the script to locate and restore the file.")
        temp = config.CUTOFF
        config.CUTOFF = util.parse_date("January 1, 2000")
        open_inquiries = mail.OpenInquiry.from_current_inbox(inbox, stat_counter.stat_labels)
        config.CUTOFF = temp
        print 'Done reading inbox for open inquiries.'

    return inbox, open_inquiries


def read_stats(stat_counter, file_base, support_mail_api):
    threads = {}

    # Obtain mail from Support Inbox and begin thread counting
    i = 0
    if config.TEST:
        file_base = "test_"+file_base
    mail_out, fmail_out, mail_writer, fmail_writer = None, None, None, None
    if config.DEBUG:
        #  Open log files
        mail_out = open('tools/logs/' + file_base + 'mail.csv', 'wb')
        fmail_out = open('tools/logs/' + file_base + 'formatted_mail.csv', 'wb')
        mail_writer = csv.writer(mail_out)
        mail_writer.writerow(['Thread ID', 'Date', 'From', 'To', 'Subject', 'X-Gmail-Labels'])
        fmail_writer = csv.writer(fmail_out)
        fmail_writer.writerow(['Thread ID', 'Date', 'From', 'To', 'Subject', 'Labels'])

    print "Reading stats label...."
    gmail_messages = googleAPI.get_messages_from_threads(support_mail_api, 'me', config.QUERY)
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
            threads[msg_id] = mail.Thread(msg, stat_counter.stat_labels)
        else:
            threads[msg_id].add_message(msg, stat_counter.stat_labels)

        if not config.COUNT_EVERY == 0 and i % config.COUNT_EVERY == 0:
            print i, msg
        i += 1

    if config.DEBUG:
        fmail_out.close()
        mail_out.close()
    print "...done"
    return threads


def evaluate_threads(threads, member_data):
    print "Evaluating threads..."
    for thread in threads:
        trd = threads[thread]

        for mem in trd.get_members():  # This isn't 100% accurate but good threads with multiple members are rare.
            member_data[mem].update_last_contact(trd.get_last_contact_date())
            if trd.is_check_in():
                member_data[mem].update_check_in(trd.get_check_in_date())

    print "...done"


def count_stats(stat_counter, threads, member_data, open_inquiries, inbox):
    print "Counting stats..."
    new_open_inquires = stat_counter.count_stats(threads, member_data)

    # num_open = Number of threads currently open excluding new open inquires
    # num_closed = Number of threads closed since the last time stats were run.
    num_open, num_closed = mail.OpenInquiry.update(open_inquiries, new_open_inquires, inbox)
    stat_counter.count_open(num_open)
    stat_counter.count_existing_closed(num_closed)
    print "...done"


def update_retention_members_only(retention_sheet_id, sheets_api, mem_stats_sheet_id, member_data, stat_labels):
    # Update member stats
    try:
        print "Updating Retention gsheet with member data..."
        new_mem_data = map(members.Member.create_stat_row, member_data.values())

        mem_request = googleAPI.update_request(mem_stats_sheet_id, new_mem_data,
                                               1, len(member_data) + 1, 0, len(stat_labels) + 3)
        sort_request = googleAPI.sort_request(mem_stats_sheet_id, 0, 1, 1000, 0, 1000)

        # If on request fails neither member or admin data wil lbe updated.
        googleAPI.spreadsheet_batch_update(sheets_api, retention_sheet_id,
                                           [mem_request, sort_request])

        print "...done"

    except HttpError, e:
        util.print_error("Error: Failed to update Member or Admin. Report and resolve error then re-run stats")
        raise e


def update_retention(retention_sheet_id, sheets_api, mem_stats_sheet_id, admin_sheet_id,
                     member_data, stat_labels, admins):
    # Update member and admin stats
    try:
        print "Updating Retention gsheet member and admin data..."
        new_mem_data = map(members.Member.create_stat_row, member_data.values())

        mem_request = googleAPI.update_request(mem_stats_sheet_id, new_mem_data,
                                               1, len(member_data) + 1, 0, len(stat_labels) + 3)
        sort_request = googleAPI.sort_request(mem_stats_sheet_id, 0, 1, 1000, 0, 1000)

        new_admin_data = map(lambda adm: members.Admin.create_stat_row(admins[adm]), sorted(admins.keys()))
        admin_request = googleAPI.update_request(admin_sheet_id, new_admin_data, 1, len(new_admin_data) + 1, 4, 6)

        # If on request fails neither member or admin data wil lbe updated.
        googleAPI.spreadsheet_batch_update(sheets_api, retention_sheet_id,
                                           [mem_request, sort_request, admin_request])

        print "...done"

    except HttpError, e:
        util.print_error("Error: Failed to update Member or Admin. Report and resolve error then re-run stats")
        raise e


def duplicate_current_tab(current_sheet_id, enrollment_dash_id, sheets_api):
    # Duplicate current tab on enrollment dash)
    new_title = util.add_days(config.END_CUTOFF, 1).strftime('%m/%d')

    try:
        print "Updating Enrollment Dashboard..."
        duplicate_request = googleAPI.duplicate_sheet_request(current_sheet_id, new_title, 3)
        googleAPI.spreadsheet_batch_update(sheets_api, enrollment_dash_id, [duplicate_request])
        print "...done"
    except HttpError:
        util.print_error("Error: Failed to duplicate current tab on Enrollment Dashboard. See steps below.")
        print '1. Duplicate Tab and rename as ' + new_title + ". If a conflict tab exists, rename the conflict tab."
        print '2. Copy all cells on duplicated tab and paste as values to remove formulas'
        print '3. Go to "Data -> Named Ranges" and remove all named ranges associated with the duplicated sheet'
        print '4. Delete all call information on the "Current" Tab of the Enrollment Dashboard'
        raw_input("Press enter to continue.")

    else:
        # Try each request individually in case one fails.
        try:
            delete_named_ranges_request = googleAPI.delete_named_range_request(
                sheets_api, enrollment_dash_id, new_title)
            googleAPI.spreadsheet_batch_update(sheets_api, enrollment_dash_id,
                                               [delete_named_ranges_request])
        except HttpError:
            util.print_error("Error: Failed to delete named ranges on duplicated current tab. See steps below.")
            print 'On the Enrollment Dashboard go to "Data -> Named Ranges" and remove all named ranges associated ' \
                  'with the tab ' + new_title
            raw_input("Press enter to continue.")
        try:
            googleAPI.remove_formulas(sheets_api, enrollment_dash_id, new_title + '!A:A')
        except HttpError:
            util.print_error("Error: Failed to remove values on duplicated current tab. See steps below.")
            print 'Copy all cells on ' + new_title + ' of the Enrollment Dashboard and paste values to remove formulas'
            raw_input("Press enter to continue.")
        try:
            googleAPI.clear_ranges(sheets_api, enrollment_dash_id, ['Current_Calls', 'Bens_Calls'])
        except HttpError:
            util.print_error("Error: Failed to clear cells on Current tab of Enrollment Dashboard. See steps below.")
            print "Delete all call information on the 'Current' Tab of the Enrollment Dashboard"
            raw_input("Press enter to continue.")


def send_stats_email(stat_counter, mail_api, to_address):
    # Draft an email
    print "Sending Stats email..."
    subject = "Stats as of " + str(util.add_days(config.END_CUTOFF, 1).strftime('%m/%d/%Y'))
    email_body = stat_counter.draft_message(config.CUTOFF, util.add_days(config.END_CUTOFF, 1))

    try:
        googleAPI.send_message(mail_api, "me", to_address, subject, email_body)
    except HttpError:
        util.print_error("Error: Failed to send email to Andy. Use text in email.txt or text printed to terminal.")
        raw_input("Press enter to continue.")
    print "...done"


def send_combined_stats_email(support_stat_counter, gov_stat_counter, mail_api, to_address):
    # Draft an email
    print "Sending Stats email..."
    subject = "Stats as of " + str(util.add_days(config.END_CUTOFF, 1).strftime('%m/%d/%Y'))

    html_body = stats.StatCounter.draft_html_message(support_stat_counter, gov_stat_counter, config.CUTOFF,
                                                     config.END_CUTOFF)

    try:
        googleAPI.send_message(mail_api, "me", to_address, subject, html_body)
    except HttpError:
        util.print_error("Error: Failed to send email to Andy. Use text in email.txt or text printed to terminal.")
        raw_input("Press enter to continue.")
    print "...done"


def gov_stats(file_base):
    member_data, stat_labels = read_members(config.GOV_MEMBER_STATS_SHEET, config.RETENTION_SPREADSHEET_ID,
                                            googleAPI.SHEETS_API, 3, config.GOV_SHORT_NAME_RANGE)
    gov_support_stat_counter = stats.StatCounter(stat_labels)
    # admins, admin_emails = read_admins(config.ADMIN_SHEET, config.SPREADSHEET_ID, googleAPI.SHEETS_API)

    if not config.SKIP:
        inbox, open_inquiries = read_inbox(gov_support_stat_counter, file_base, googleAPI.GOV_SUPPORT_MAIL_API)
        threads = read_stats(gov_support_stat_counter, file_base, googleAPI.GOV_SUPPORT_MAIL_API)
        # update_admin_dates(threads, admins, admin_emails)
        evaluate_threads(threads, member_data)

        count_stats(gov_support_stat_counter, threads, member_data, open_inquiries, inbox)

    print "Updating Weekly Support Stats gsheet..."
    # Combine and format stats in prep for writing.
    #   Combines changes requests, CITI, Issues, Sales Ping. Calculates totals
    gov_support_stat_counter.format_stats()
    # Update weekly support stats gsheet
    gov_support_stat_counter.update_weekly_support_stats(googleAPI.SHEETS_API, config.WEEKLY_STATS_SPREADSHEET_ID,
                                                         config.GOV_WEEKLY_STATS_SHEET_ID)
    print "...done"

    # Update Google Sheets
    update_retention_members_only(config.RETENTION_SPREADSHEET_ID, googleAPI.SHEETS_API,
                                  config.GOV_MEMBER_STATS_SHEET_ID,
                                  member_data, stat_labels)
    return gov_support_stat_counter


def support_stats():
    member_data, stat_labels = read_members(config.MEMBER_STATS_SHEET, config.RETENTION_SPREADSHEET_ID,
                                            googleAPI.SHEETS_API, 3, config.SHORT_NAME_RANGE)
    support_stat_counter = stats.StatCounter(stat_labels)
    admins, admin_emails = read_admins(config.ADMIN_SHEET, config.RETENTION_SPREADSHEET_ID, googleAPI.SHEETS_API)

    if not config.SKIP:
        inbox, open_inquiries = read_inbox(support_stat_counter, "", googleAPI.SUPPORT_MAIL_API)
        threads = read_stats(support_stat_counter, "", googleAPI.SUPPORT_MAIL_API)
        update_admin_dates(threads, admins, admin_emails)
        evaluate_threads(threads, member_data)

        count_stats(support_stat_counter, threads, member_data, open_inquiries, inbox)

    print "Updating Weekly Support Stats gsheet..."
    # Combine and format stats in prep for writing.
    #   Combines changes requests, CITI, Issues, Sales Ping. Calculates totals
    support_stat_counter.format_stats()
    support_stat_counter.get_support_calls()
    # Update weekly support stats gsheet
    support_stat_counter.update_weekly_support_stats(googleAPI.SHEETS_API, config.WEEKLY_STATS_SPREADSHEET_ID)
    print "...done"

    # Update Google Sheets
    update_retention(config.RETENTION_SPREADSHEET_ID, googleAPI.SHEETS_API, config.MEMBER_STATS_SHEET_ID,
                     config.ADMIN_SHEET_ID,
                     member_data, stat_labels, admins)
    duplicate_current_tab(config.CURRENT_SHEET_ID, config.ENROLLMENT_DASHBOARD_ID, googleAPI.SHEETS_API)
    return support_stat_counter


if __name__ == '__main__':
    # TODO support gov stats only
    # TODO add support for more than one additional inbox
    # TODO Perform gsheet updates after all stats gathered.
    # TODO Preface print/error statements with "Gov" or "Support"
    # TODO Move all hardcoded values / global variables to config or create init method
    # TODO Rebuild application shortcut. Allow for shortcut to run rom desktop
    support_counter = support_stats()
    # send_stats_email(support_counter, googleAPI.MAIL_API, config.STATS_TO_ADDRESS)
    stats.Stat._id = 0
    gov_counter = gov_stats("gov_")
    send_combined_stats_email(support_counter, gov_counter, googleAPI.MAIL_API, config.STATS_TO_ADDRESS)
