import time
import csv
import datetime

from googleapiclient.errors import HttpError
import googleAPI
import mail
from tools import config
import util


class Stat:
    """
    Stat has a name, priority and a count. The first initialized Stat is assigned priority 0 unless otherwise specified.
    Priority is incremented by one for each subsequent Stat unless otherwise specified.
    """
    _id = 0  # First stat will have priority 0. This makes indexing easier when counting.

    def __init__(self, name, count=0, priority=-99):
        """
        Constructs a new Stat
        :param name: stat name
        :param count: stat count (default = 0)
        :param priority: stat priority (default = -99: Will be assigned next available priority)
        """
        if priority == -99:
            priority = Stat._id
            Stat._id += 1

        self.stat = name
        self._count = count
        self._priority = priority

    def increment(self, i=1):
        """
        Increments stat count by i
        :param i: (default=1)
        :return:
        """
        self._count += i

    def decrement(self, i=1):
        """
        Decrements stat count by i
        :param i: (default=1)
        :return:
                """
        self._count -= i

    def set_count(self, i):
        """
        Sets the stat count to i
        :param i: New stat count.
        :return:
        """
        self._count = i

    def get_count(self):
        """
        Returns current stat count
        :return:
        """
        return self._count

    def get_priority(self):
        """
        Returns current priority
        :return:
        """
        return self._priority

    def set_priority(self, p):
        """
        Sets the priority to new value
        :param p: New priority
        :return:
        """

        self._priority = p

    def __str__(self):
        return str(self.stat) + ": (" + str(self._count) + ", " + str(self._priority) + ")"

    def __repr__(self):
        return str(self.stat) + ": (" + str(self._count) + ", " + str(self._priority) + ")"

    def __eq__(self, other):
        return isinstance(other, Stat) and self._priority == other._priority

    def __le__(self, other):
        return isinstance(other, Stat) and self._priority <= other._priority

    def __lt__(self, other):
        return isinstance(other, Stat) and self._priority < other._priority

    def __ge__(self, other):
        return isinstance(other, Stat) and self._priority >= other._priority

    def __gt__(self, other):
        return isinstance(other, Stat) and self._priority > other._priority


PINGS = ["Sales Pings", "User Inquiries", "Demo Requests", "New Organizations", "Voicemails", "Total Pings", "Web Form"]

TOTALS = ["Overall Total New Inquiries", "New Open Inquiries", "New Closed Inquiries", "Existing Open Inquiries Closed",
          "Total Open Inquiries", "Total Closed Inquiries"]

CALL_INFO = ["Sessions", "Sales Calls", "Demos", "Institutions"]


def from_list(lst):
    """
    Initializes a new dictionary of Stat objects from the provided list.
    Stat objects will begin at priority 0 and have a count 0
    :param lst: List of stat names.
    :return: created dictionary
    """
    result = {}
    i = 0
    for label in lst:
        result[label] = Stat(label, 0, i)
        i += 1

    return result


PING_STATS = from_list(PINGS)
TOTAL_STATS = from_list(TOTALS)
CALL_STATS = from_list(CALL_INFO)

STAT_LABELS = {}


def count_user_inquiry(i=1):
    try:
        PING_STATS["User Inquiries"].increment(i)
    except KeyError:
        raise KeyError("Could not increment Inquiry")


def count_demo(i=1):
    try:
        PING_STATS["Demo Requests"].increment(i)
    except KeyError:
        raise KeyError("Could not increment Demo Request")


def count_new_org(i=1):
    try:
        PING_STATS["New Organizations"].increment(i)
    except KeyError:
        raise KeyError("Could not increment New Orgs")


def count_voicemail(i=1):
    try:
        PING_STATS["Voicemails"].increment(i)
    except KeyError:
        raise KeyError("Could not increment VM")


def count_sales_ping(i=1):
    try:
        PING_STATS["Sales Pings"].increment(i)
    except KeyError:
        raise KeyError("Could not increment Sales Ping")


def count_web_form(i=1):
    try:
        PING_STATS["Web Form"].increment(i)
    except KeyError:
        raise KeyError("Could not increment Web Form")


def count_open(i=1):
    try:
        TOTAL_STATS["Total Open Inquiries"].increment(i)
    except KeyError:
        raise KeyError("Could not increment Total Open")


def count_new_open(i=1):
    try:
        TOTAL_STATS["New Open Inquiries"].increment(i)
    except KeyError:
        raise KeyError("Could not increment New Open")


def count_new_closed(i=1):
    try:
        TOTAL_STATS["New Closed Inquiries"].increment(i)
    except KeyError:
        raise KeyError("Could not increment New Closed")


def count_existing_closed(i=1):
    try:
        TOTAL_STATS["Existing Open Inquiries Closed"].increment(i)
    except KeyError:
        raise KeyError("Could not increment Existing Open Inquiries Closed")


def calc_total_closed():
    count = TOTAL_STATS["New Closed Inquiries"].get_count() + TOTAL_STATS["Existing Open Inquiries Closed"].get_count()
    TOTAL_STATS["Total Closed Inquiries"].increment(count)
    return TOTAL_STATS["Total Closed Inquiries"]


def format_stats():
    """
    Combines and formats stat dictionaries in preparation for writing.
    'Change Request' += 'Change Request - Access Level'
    'Issues' = 'Issue' + 'System Access Issue' + 'Issue/PDF'
    'CITI' = 'CITI Integration' + 'CITI Interface Errors'

    Moves 'Welcome to Support', 'Welcome Ping', and Sales to PING_STATS
    Combines Sales Pings and Web form resulting in 'Sales Pings(Web Form)' as the new count for 'Sales Pings'

    Sums ping and non-ping totals

    Updates CALL_STATS with support call information.

    :return:
    """
    # Combine non-pings
    change = STAT_LABELS["Change Request"].get_count()
    access = STAT_LABELS["Change Request - Access Level"].get_count()
    STAT_LABELS["Change Request"].set_count(change + access)
    del STAT_LABELS["Change Request - Access Level"]

    issue = STAT_LABELS["Issue"].get_count()
    system = STAT_LABELS["System Access Issue"].get_count()
    pdf = STAT_LABELS["Issue/PDF"].get_count()

    issue_priority = STAT_LABELS["Issue"].get_priority()
    STAT_LABELS["Issues"] = Stat("Issues", issue + system + pdf, issue_priority)
    del STAT_LABELS["Issue"]
    del STAT_LABELS["System Access Issue"]
    del STAT_LABELS["Issue/PDF"]

    citi1 = STAT_LABELS["CITI Integration"].get_count()
    citi_priority = STAT_LABELS["CITI Integration"].get_priority()
    citi2 = STAT_LABELS["CITI Interface Errors"].get_count()
    del STAT_LABELS["CITI Integration"]
    del STAT_LABELS["CITI Interface Errors"]
    STAT_LABELS["CITI"] = Stat("CITI", citi1 + citi2, citi_priority)
    assert "CITI" in STAT_LABELS

    pings_in_stats = ["Welcome to Support", "Welcome Ping", "Sales"]
    i = PING_STATS["Sales Pings"].get_priority() - len(pings_in_stats)
    for ping in pings_in_stats:
        PING_STATS[ping] = STAT_LABELS[ping]
        del STAT_LABELS[ping]
        PING_STATS[ping].set_priority(i)
        i += 1

    # Calculate total Pings and Non-Pings
    total_non_pings = 0
    for stat in STAT_LABELS:
        total_non_pings += STAT_LABELS[stat].get_count()

    STAT_LABELS["Total Non-Pings"] = Stat("Total Non-Pings", total_non_pings, issue_priority + 0.1)

    total_pings = PING_STATS["User Inquiries"].get_count() + PING_STATS["Demo Requests"].get_count() + \
        PING_STATS["Welcome to Support"].get_count() + PING_STATS["Welcome Ping"].get_count() + \
        PING_STATS["Sales"].get_count() + PING_STATS["New Organizations"].get_count() + \
        PING_STATS["Sales Pings"].get_count() + \
        PING_STATS["Voicemails"].get_count()

    # Combine pings
    sp = PING_STATS["Sales Pings"].get_count()
    web = PING_STATS["Web Form"].get_count()
    PING_STATS["Sales Pings"].set_count(str(sp) + "(" + str(web) + ")")
    del PING_STATS["Web Form"]

    PING_STATS["Total Pings"].set_count(total_pings)

    # Calculate total open and closed
    TOTAL_STATS["Overall Total New Inquiries"].set_count(total_pings + total_non_pings)
    calc_total_closed()
    TOTAL_STATS["Total Open Inquiries"].increment(TOTAL_STATS['New Open Inquiries'].get_count())
    # Get call info
    get_support_calls()


def extract_labels(data):
    """
    Creates a new Stat object for each item in data and adds it to the STAT_LABELS dictionary.
    :param data: List of stats.
    :return: {stat: Stat(stat)}
    """
    for stat in data:
        STAT_LABELS[stat] = Stat(stat)

    return STAT_LABELS


def _sort_stats(stats, dictionary):
    """
    Sorts a list of stats by assigned priority from lo to hi.
    :param stats: List of stat labels to be sorted.
    :param dictionary: Dictionary mapping stat to Stat object.
    :return: A list of sorted statistics from lo to hi using on stat priority in dictionary.
    """
    i = 1
    while i < len(stats):
        j = i
        while j > 0 and dictionary[stats[j-1]].get_priority() > dictionary[stats[j]].get_priority():
            swap = stats[j-1]
            stats[j-1] = stats[j]
            stats[j] = swap
            j = j - 1
        i = i + 1

    return stats


def _write_stat_row(writer, thread, stat):
    """
    Writes basic thread information and the to the specified csv writer. Used to create a log file.
    [Oldest Date, Subject, Stat Labels, Member Labels, Counted Statistic, Thread Closed?]
    :param writer: CSV Writer used to write the row.
    :param thread: Thread object providing thread information
    :param stat: Statistic that was be counted for this thread.
    :return:
    """
    writer.writerow([thread.get_oldest_date(), thread.get_subject(), thread.get_stats(), thread.get_members(),
                    stat, thread.is_closed()])


def count_stats(threads, members):
    """
    Counts the number of statistics and updates the STAT_LABELS, PING_STATS, TOTAL_STATS dictionaries.
    :param threads: Threads for which stats will be determined.
    :param members: Members whos stats will be updated.
    :return: Dictionary of new open inquiries.
    """
    out = None
    writer = None
    if config.DEBUG:
        try:
            out = open('tools/logs/thread_lookup.csv', 'wb')
            writer = csv.writer(out)
            writer.writerow(["Date", "Subject", "Stats", "Members", "Counted Stat", "Closed?"])
        except IOError:
            util.print_error("Error. Could not open thread_lookup. Results will not be logged.")

    open_inquiries = {}
    for thread in threads:
        trd = threads[thread]
        if trd.is_good():
            if not trd.is_non_ping():
                #  Handle Pings
                if trd.is_inquiry():
                    if trd.is_sales_ping():
                        count_web_form()
                        count_sales_ping()
                        if out is not None:
                            _write_stat_row(writer, trd, "Sales Pings")
                    else:
                        count_user_inquiry()
                        if out is not None:
                            _write_stat_row(writer, trd, "Inquiry")
                elif trd.is_demo():
                    if trd.is_sales_ping():
                        count_web_form()
                        count_sales_ping()
                        if out is not None:
                            _write_stat_row(writer, trd, "Sales Pings")
                    else:
                        if out is not None:
                            _write_stat_row(writer, trd, "Demo")
                        count_demo()
                elif trd.is_sales_ping():
                    count_sales_ping()
                elif trd.is_vm():
                    count_voicemail(trd.get_count())
                    if out is not None:
                        _write_stat_row(writer, trd, "Voicemail")
                elif trd.is_new_org():
                    count_new_org()
                    if out is not None:
                        _write_stat_row(writer, trd, "New Org")
            else:
                # Handle Non-pings, Count lowest priority stat only
                sorted_stats = _sort_stats(trd.get_stats(), STAT_LABELS)
                if len(sorted_stats) > 0:
                    counted_stat = sorted_stats[0]
                    STAT_LABELS[counted_stat].increment()
                    # Sales inquiries are considered Pings
                    if counted_stat not in ["Sales"]:
                        if trd.is_closed():
                            count_new_closed()
                        else:
                            count_new_open()
                            open_inquiries[trd.get_id()] = mail.OpenInquiry(trd.get_id(), trd.get_subject())
                        if out is not None:
                            _write_stat_row(writer, trd, counted_stat)

                    for mem in trd.get_members():
                        for stat in sorted_stats:
                            members[mem].increment_stat(STAT_LABELS[stat].get_priority())
                else:
                    # TODO Error Log
                    pass
    if out is not None:
        out.close()
    return open_inquiries


def get_retention_calls():
    """
    Attempts to obtain the number of retention calls from teh Retention spreadsheet. If this attempt fails the
    user is prompted to enter the number of calls.
    :return: The number of retention calls
    """

    try:
        retention_calls = googleAPI.SHEETS_API.spreadsheets().values().get(
            spreadsheetId=config.SPREADSHEET_ID, range='Weekly_Check_In')\
            .execute().get('values', [])
        return str(retention_calls[0][0])
    except HttpError:
        util.print_error("Error: Failed to read number of check in calls.")
        retention_calls = raw_input(
            "Enter the number of check in calls this week (cell B2 on Chart Data tab of Retention sheet).   ")
        return retention_calls


def _enter_calls():
    """
    Asks the user to enter the number of Support Calls
    :return: number of sessions, sales calls, sales demos, demo institutions
    """
    sessions = raw_input("\n\nEnter the Total # of Sessions...  ")
    sales_calls = raw_input("Enter the Total # of Sales Calls...  ")
    sales_demos = raw_input("Enter the Total # of Sales Demos...  ")
    demo_institutions = raw_input("Enter Sales Demo institutions...  ")
    return sessions, sales_calls, sales_demos, demo_institutions


def get_support_calls():
    """
    Attempts to obtain the number of support calls from the Enrollment Dashboard. Updates the CALL_STATS with correct
    values.
    :return: None.
    """
    try:
        calls = googleAPI.get_range('Call_Info', config.ENROLLMENT_DASHBOARD_ID, googleAPI.SHEETS_API, 'COLUMNS')
    except HttpError, e:
        print e
        util.print_error("Error: Failed to read call info from Enrollment Dashboard. Please report.")
        sessions, sales_calls, sales_demos, demo_institutions = _enter_calls()
    else:
        if len(calls) < 3:
            sessions, sales_calls, sales_demos, demo_institutions = _enter_calls()
        else:
            sessions = calls[0][0]
            sales_calls = calls[1][0]
            sales_demos = calls[2][0]
            if len(calls) > 3:
                demo_institutions = str(calls[3][0])
            else:
                demo_institutions = ""

    CALL_STATS["Sessions"].set_count(sessions)
    CALL_STATS["Sales Calls"].set_count(sales_calls)
    CALL_STATS["Demos"].set_count(sales_demos)
    CALL_STATS["Institutions"].set_count(demo_institutions)


def draft_message(cutoff):
    """
    Writes an email summarizing weekly statistics.
    :param cutoff: Earliest date for which stats should be counted.
    :return: the email text
    """
    initials = raw_input("Enter your initials...  ")

    today = time.strftime('%m/%d/%Y')
    today_weekday = time.strftime('%a')
    start_date = cutoff.strftime('%m/%d/%Y')
    start_weekday = cutoff.strftime('%a')

    total_pings = PING_STATS["Total Pings"].get_count()
    total_non_pings = STAT_LABELS["Total Non-Pings"].get_count()
    sp = PING_STATS["Sales Pings"].get_count()
    pings_less_sales = str(total_pings - int(sp[:str(sp).find("(")]))  # This removes web forms.
    total = str(total_pings + total_non_pings)

    txt = "Andy, \n\n " \
        "Below are the requested statistics from " + str(start_weekday) + ", " + start_date + " to " + \
          str(today_weekday) + ", " + str(today) + ":\n\n" \
        "Pings (includes New Orgs; no Sales Pings): " + pings_less_sales + "\n" \
        "Non-Pings: " + str(total_non_pings) + "\n" \
        "Sales Inquiries: " + str(PING_STATS["Sales Pings"].get_count()) + "\n" \
        "Overall Total New Inquiries: " + total + "\n\n" \
        "Total New Inquiries (Non-Pings) Currently Open: " + str(TOTAL_STATS["New Open Inquiries"].get_count()) + "\n" \
        "Total New Inquiries (Non-Pings) Closed: " + str(TOTAL_STATS["New Closed Inquiries"].get_count()) + "\n" \
        "Total Existing Open Inquiries Closed: " + str(TOTAL_STATS["Existing Open Inquiries Closed"].get_count()) + \
          "\nTotal Open Inquiries: " + str(TOTAL_STATS["Total Open Inquiries"].get_count()) + "\n" \
        "Total Closed Inquiries: " + str(TOTAL_STATS["Total Closed Inquiries"].get_count()) + "\n\n" \
        "Total # of Sessions: " + str(CALL_STATS["Sessions"].get_count()) + "\n" \
        "Total # of Sales Calls: " + str(CALL_STATS["Sales Calls"].get_count()) + "\n" \
        "Total # of Demo Calls: " + str(CALL_STATS["Demos"].get_count()) + " " + \
          str(CALL_STATS["Institutions"].get_count()) + "\n" \
        "Total # of Retention Calls: " + str(get_retention_calls()) + "\n\n" \
        "Cumulative Weekly Statistics can be accessed via the following sheet:\n" \
        "https://docs.google.com/a/irbnet.org/spreadsheets/d/" + config.WEEKLY_STATS_SHEET_ID + "\n\n" \
        "Retention information can be accessed via the following sheet\n" \
        "https://docs.google.com/spreadsheets/d/" + config.SPREADSHEET_ID + "\n\n" \
        "Let me know if you have any questions. Thanks!\n\n" + initials

    return txt


def add_stats(dictionary):
    """
    Creates a list of stat counts sorted by the respective stat priority for the given stat dictionary.
    :param dictionary: {name : Stat} dictionary whose stats will be sorted.
    :return: A list of stat counts sorted by stat priority lo -> hi
    """
    values = []
    for item in _sort_stats(dictionary.keys(), dictionary):
        if type(dictionary[item].get_count()) is str:
            values.append(((dictionary[item].get_count()), 'STRING'))
        else:
            values.append(((dictionary[item].get_count()), 'NUMBER'))

    return values


def update_weekly_support_stats(service, spreadsheet_id, sheet_id=0):
    """
    Updates the specified sheet with aggregated statistics info by inserting a new column between columns B and C.
    The new column has the format.
        Date

        STAT_LABELS


        PING_STATS


        TOTAL_STATS

        CALL_STATS

    :param service: Authorized Google Sheets service to access the Sheets APU
    :param spreadsheet_id: Spreadsheet ID for the sheet that will be updated.
    :param sheet_id: Google Sheet ID for the sheet that will be updated.
    :return: None
    """

    values = [(util.serial_date(datetime.datetime.today()), 'DATE')]
    values.extend(add_stats(STAT_LABELS))
    values.extend([("", 'STRING'), ("", 'STRING')])
    values.extend(add_stats(PING_STATS))
    values.extend([("", 'STRING'), ("", 'STRING')])
    values.extend(add_stats(TOTAL_STATS))
    values.append(("", 'STRING'))
    values.extend(add_stats(CALL_STATS))
    insert_column_request = googleAPI.insert_column_request(sheet_id, values, 0, 39, 2, 3)
    try:
        googleAPI.spreadsheet_batch_update(service, spreadsheet_id, insert_column_request)
    except HttpError:
        util.print_error("Failed to update Weekly Support Stats. ")
        try:
            out_file = open('stats.csv', 'w')
            out = csv.writer(out_file)
            for (v, t) in values:
                out.writerow([v])
            print "Use stats.csv to update Weekly Support Stats Sheet"
            out_file.close()
        except IOError:
            util.print_error("Failed to write stats.csv. See stats below.")
            for (v, t) in values:
                print v
