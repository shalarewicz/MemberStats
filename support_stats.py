import csv
from tools import config
from tools.lib import mail, members, stats, util, googleAPI
from googleapiclient.errors import HttpError

MAIL_API, SHEETS_API, SUPPORT_MAIL_API, GOV_SUPPORT_MAIL_API = None, None, None, None


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
    inbox_threads = googleAPI.get_thread_ids(support_mail_api, "me", "label:Inbox")
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
        temp = START_DATE
        config.START_DATE = util.parse_date("January 1, 2000")
        inbox = googleAPI.get_messages_from_threads(support_mail_api, "me", "label:Inbox")
        open_inquiries = mail.OpenInquiry.from_current_inbox(inbox, stat_counter.stat_labels)
        config.START_DATE = temp
        print 'Done reading inbox for open inquiries.'

    return inbox_threads, open_inquiries


def read_stats(stat_counter, file_base, support_mail_api, cutoff, member_labels):
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
            threads[msg_id] = mail.Thread(msg, stat_counter.stat_labels, cutoff, member_labels)
        else:
            threads[msg_id].add_message(msg, stat_counter.stat_labels, cutoff, member_labels)

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


def count_stats(stat_counter, threads, member_data, open_inquiries, inbox, file_base=""):
    print "Counting stats..."
    new_open_inquires = stat_counter.count_stats(threads, member_data, file_base)

    # num_open = Number of threads currently open excluding new open inquires
    # num_closed = Number of threads closed since the last time stats were run.
    num_open, num_closed, updated_open_inquiries = mail.OpenInquiry.update(open_inquiries, new_open_inquires, inbox)
    stat_counter.count_open(num_open)
    stat_counter.count_existing_closed(num_closed)
    print "...done"
    return updated_open_inquiries


def get_retention_member_update_requests(mem_stats_sheet_id, member_data, stat_labels):
    """
    Gets requests to update member data
    :param mem_stats_sheet_id: Google Sheet ID specifying the destination sheet for updated member data
    :param member_data: New member data must contain a value for each stat label
    :param stat_labels: Stat Labels used as header on specified shet
    :return: [Member Update Request, Member Sort Request]
    """
    new_mem_data = map(members.Member.create_stat_row, member_data.values())

    mem_request = googleAPI.update_request(mem_stats_sheet_id, new_mem_data,
                                           1, len(member_data) + 1, 0, len(stat_labels) + 3)
    sort_request = googleAPI.sort_request(mem_stats_sheet_id, 0, 1, 1000, 0, 1000)

    return [mem_request, sort_request]


def get_retention_admin_update_requests(admin_sheet_id, admins):
    """
    Gets requests to update admin data
    :param admin_sheet_id: Google Sheet ID specifying the destination sheet for updated admin data
    :param admins: new member data
    :return: Admin Update Request
    """
    new_admin_data = map(lambda adm: members.Admin.create_stat_row(admins[adm]), sorted(admins.keys()))
    return googleAPI.update_request(admin_sheet_id, new_admin_data, 1, len(new_admin_data) + 1, 4, 6)


def update_retention(sheets_api, retention_sheet_id, requests):
    """
    Updates the specified retention sheet with all listed requests.
    :param sheets_api: Authorized Google Sheets service
    :param retention_sheet_id: Spreadsheet ID of the Retention sheet being updated
    :param requests: List of Google batch update requests
    :return:
    """

    try:
        print "Updating Retention gsheet"
        googleAPI.spreadsheet_batch_update(sheets_api, retention_sheet_id, requests)
        print "...done"

    except HttpError, e:
        # todo print stack trace to a log?
        util.print_error("Error: Failed to update Retention Sheet No Changes have been made. "
                         "Report and resolve error then re-run stats")
        raise e


def duplicate_current_tab(sheets_api, current_sheet_id, enrollment_dash_id):
    # Duplicate current tab on enrollment dash)
    new_title = END_DATE.strftime('%m/%d')

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


def send_stats_email(stat_counter, mail_api, to_address, subject):
    # Draft an email
    print "Sending Stats email..."
    email_body = stat_counter.draft_message(START_DATE, END_DATE,
                                            stats.get_retention_calls(
                                                SHEETS_API, config.RETENTION_SPREADSHEET_ID))

    try:
        googleAPI.send_message(mail_api, "me", to_address, subject, email_body)

    except HttpError:
        util.print_error("Error: Failed to send email to Andy. Use text in email.txt or text printed to terminal.")
        raw_input("Press enter to continue.")
    print "...done"


def send_combined_stats_email(support_stat_counter, gov_stat_counter, mail_api, to_address):
    # Draft an email
    print "Sending Stats email..."
    subject = "Stats as of " + str(END_DATE.strftime('%m/%d/%Y'))

    html_body = stats.StatCounter.draft_html_message(support_stat_counter, gov_stat_counter, START_DATE,
                                                     END_DATE,
                                                     stats.get_retention_calls(
                                                         SHEETS_API, config.RETENTION_SPREADSHEET_ID))

    try:
        googleAPI.send_message(mail_api, "me", to_address, subject, html_body)
    except HttpError:
        util.print_error("Error: Failed to send email to Andy. Use text in email.txt or text printed to terminal.")
        raw_input("Press enter to continue.")
    print "...done"


def run_stats():
    # TODO add support for more than one additional inbox
    # TODO Rebuild application shortcut. Allow for shortcut to run rom desktop
    support_counter, support_open_inquiries = None, None
    update_requests = []
    if not config.GOV:
        print "Running SUPPORT stats...."
        support_counter, update_requests, support_open_inquiries = \
            get_stats(SUPPORT_MAIL_API, SHEETS_API, START_DATE, "", config.MEMBER_STATS_SHEET,
                      config.MEMBER_STATS_SHEET_ID,
                      config.SHORT_NAME_RANGE,
                      config.RETENTION_SPREADSHEET_ID,
                      config.ADMIN_SHEET, config.ADMIN_SHEET_ID,
                      True, True)

    # Reset id to for ay new counter
    stats.Stat._id = 0
    print "Running GOV stats...."
    gov_counter, gov_update_requests, gov_open_inquiries = \
        get_stats(GOV_SUPPORT_MAIL_API, SHEETS_API, START_DATE, "gov_",
                  config.GOV_MEMBER_STATS_SHEET,
                  config.GOV_MEMBER_STATS_SHEET_ID,
                  config.GOV_SHORT_NAME_RANGE,
                  config.RETENTION_SPREADSHEET_ID,
                  None, None, with_admins=False,  # No admin sheets for gov
                  with_support_calls=False)
    update_requests.extend(gov_update_requests)

    print "Updating Weekly Support Stats gsheet..."
    if not config.GOV:
        support_counter.update_weekly_support_stats(SHEETS_API, config.WEEKLY_STATS_SPREADSHEET_ID)
    gov_counter.update_weekly_support_stats(SHEETS_API, config.WEEKLY_STATS_SPREADSHEET_ID,
                                            config.GOV_WEEKLY_STATS_SHEET_ID)
    print "...done"

    print "Print Updating Retention Sheet and Enrollment Dashboard"
    update_retention(SHEETS_API, config.RETENTION_SPREADSHEET_ID, update_requests)
    if not config.GOV:
        duplicate_current_tab(SHEETS_API, config.CURRENT_SHEET_ID, config.ENROLLMENT_DASHBOARD_ID)
    print "...done"

    if not config.SKIP:
        if not config.GOV:
            mail.OpenInquiry.write_to_file(support_open_inquiries.values(), 'open.txt')

        mail.OpenInquiry.write_to_file(gov_open_inquiries.values(), 'gov_open.txt')

    if config.GOV:
        subject = "Gov Stats as of " + END_DATE.strftime('%m/%d/%Y')
        send_stats_email(gov_counter, MAIL_API, config.STATS_TO_ADDRESS, subject)
    else:
        send_combined_stats_email(support_counter, gov_counter, MAIL_API, config.STATS_TO_ADDRESS)


def get_stats(support_mail_api, sheets_api, cutoff,
              file_base,
              member_stats_sheet,
              member_stats_sheet_id,
              short_name_range,
              retention_spreadsheet_id,
              admin_sheet,
              admin_sheet_id,
              with_admins=True,
              with_support_calls=True):
    """

    :param file_base:
    :param cutoff
    :param member_stats_sheet:
    :param member_stats_sheet_id:
    :param short_name_range:
    :param admin_sheet:
    :param admin_sheet_id:
    :param retention_spreadsheet_id:
    :param support_mail_api:
    :param sheets_api:
    :param with_admins:
    :param with_support_calls:
    :return: StatCounter, List of Google Sheet Update Requests, dictionary of open_inquiries
    """

    member_data, stat_labels = read_members(member_stats_sheet, retention_spreadsheet_id, sheets_api, 3,
                                            short_name_range)
    support_stat_counter = stats.StatCounter(stat_labels)

    updated_open_inquiries = None
    update_requests = []
    if not config.SKIP:
        inbox, open_inquiries = read_inbox(support_stat_counter, file_base, support_mail_api)
        threads = read_stats(support_stat_counter, file_base, support_mail_api, cutoff, member_data.keys())
        if with_admins:
            admins, admin_emails = read_admins(admin_sheet, retention_spreadsheet_id, sheets_api)
            update_admin_dates(threads, admins, admin_emails)
            admin_update_request = get_retention_admin_update_requests(admin_sheet_id, admins)
            update_requests.append(admin_update_request)

        evaluate_threads(threads, member_data)

        updated_open_inquiries = count_stats(support_stat_counter, threads, member_data, open_inquiries,
                                             inbox, file_base)

    # Combine and format stats in prep for writing.
    #   Combines changes requests, CITI, Issues, Sales Ping. Calculates totals
    support_stat_counter.format_stats()

    if with_support_calls:
        support_stat_counter.get_support_calls(SHEETS_API)

    # Update Google Sheets
    mem_update_request = get_retention_member_update_requests(member_stats_sheet_id, member_data, stat_labels)
    update_requests.extend(mem_update_request)

    return support_stat_counter, update_requests, updated_open_inquiries


def init_param():
    util.parse()
    util.set_test(config.TEST)
    util.print_param()
    global START_DATE, END_DATE
    START_DATE = util.get_cutoff_date(
        "\nOn which date were stats last run?\ni.e. "
        "What is the earliest date for which stats should count, typically last Thursday?\n")

    # This date is inclusive of when stats should count.
    END_DATE = util.get_cutoff_date(
        "\nEnter the final date for which stats should count. Typically this Wednesday.\n")
    END_DATE = util.add_days(END_DATE, 1)

    if END_DATE < START_DATE:
        raise RuntimeError("Start Date must be less than or equal to End Date.")

    # Update query to use cutoff dates
    config.QUERY = "after:" + START_DATE.strftime('%Y/%m/%d') + " before:" + \
                   END_DATE.strftime('%Y/%m/%d') + " " + config.QUERY

    # Create google api service objects
    global MAIL_API, SHEETS_API, SUPPORT_MAIL_API, GOV_SUPPORT_MAIL_API
    MAIL_API = googleAPI.get_api('gmail', 'v1', 'personal', googleAPI.SCOPES, 3)
    SHEETS_API = googleAPI.get_api('sheets', 'v4', 'personal', googleAPI.SCOPES, 3)
    if not config.GOV:
        SUPPORT_MAIL_API = googleAPI.get_api('gmail', 'v1', 'support', googleAPI.SUPPORT_SCOPE, 3)
    GOV_SUPPORT_MAIL_API = googleAPI.get_api('gmail', 'v1', 'gov_support', googleAPI.GOV_SUPPORT_SCOPE, 3)


if __name__ == '__main__':
    START_DATE = util.parse_date("January 1, 2000")
    END_DATE = util.parse_date("January 1, 2000")
    init_param()
    run_stats()
