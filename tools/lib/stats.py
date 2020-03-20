import csv
import datetime

from googleapiclient.errors import HttpError
import googleAPI
import mail
from tools import config
import util
from email.mime.text import MIMEText
# TODO pip install email. See if it's possible to create local copies


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


#  TODO Move these to the config file
PINGS = ["Sales Pings", "User Inquiries", "Demo Requests", "Support Pings",
         "New Organizations", "Voicemails", "Total Pings", "Web Form"]

TOTALS = ["Overall Total New Inquiries", "New Open Inquiries", "New Closed Inquiries", "Existing Open Inquiries Closed",
          "Total Open Inquiries", "Total Closed Inquiries"]

CALL_INFO = ["Sessions", "Sales Calls", "Demos", "Institutions"]

VOICEMAILS = [config.VM_RESEARCHER, config.VM_ADMIN, config.VM_SALES, config.VM_FINANCE]


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


def get_retention_calls():
    """
    Attempts to obtain the number of retention calls from teh Retention spreadsheet. If this attempt fails the
    user is prompted to enter the number of calls.
    :return: The number of retention calls
    """

    try:
        retention_calls = googleAPI.SHEETS_API.spreadsheets().values().get(
            spreadsheetId=config.RETENTION_SPREADSHEET_ID, range='Weekly_Check_In')\
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


def extract_labels(data):
    """
    Creates a new Stat object for each item in data and adds it to the STAT_LABELS dictionary.
    :param data: List of stats.
    :return: {stat: Stat(stat)}
    """
    result = {}
    for stat in data:
        result[stat] = Stat(stat)

    return result


class StatCounter:
    def __init__(self, labels):
        self.ping_stats = from_list(PINGS)
        self.total_stats = from_list(TOTALS)
        self.call_stats = from_list(CALL_INFO)
        self.vm_stats = from_list(VOICEMAILS)

        self.stat_labels = extract_labels(labels)

    def set_labels(self, stats_labels):
        self.stat_labels = extract_labels(stats_labels)
        print "\nThe following Statistics will be determined..."
        for stat in sorted(self.stat_labels):
            print stat

    def count_user_inquiry(self, i=1):
        try:
            self.ping_stats["User Inquiries"].increment(i)
        except KeyError:
            raise KeyError("Could not increment Inquiry")

    def count_demo(self, i=1):
        try:
            self.ping_stats["Demo Requests"].increment(i)
        except KeyError:
            raise KeyError("Could not increment Demo Request")

    def count_support_ping(self, i=1):
        try:
            self.ping_stats["Support Pings"].increment(i)
        except KeyError:
            raise KeyError("Could not increment Support Pings")

    def count_new_org(self, i=1):
        try:
            self.ping_stats["New Organizations"].increment(i)
        except KeyError:
            raise KeyError("Could not increment New Orgs")

    def count_admin_vm(self, i=1):
        try:
            self.vm_stats[config.VM_ADMIN].increment(i)
        except KeyError:
            raise KeyError("Could not increment Admin VM")

    def count_sales_vm(self, i=1):
        try:
            self.vm_stats[config.VM_SALES].increment(i)
        except KeyError:
            raise KeyError("Could not increment Admin VM")

    def count_finance_vm(self, i=1):
        try:
            self.vm_stats[config.VM_FINANCE].increment(i)
        except KeyError:
            raise KeyError("Could not increment Admin VM")

    def count_res_vm(self, i=1):
        try:
            self.vm_stats[config.VM_RESEARCHER].increment(i)
            self.ping_stats["Voicemails"].increment(i)
        except KeyError:
            raise KeyError("Could not increment Researcher VM")

    def count_sales_ping(self, i=1):
        try:
            self.ping_stats["Sales Pings"].increment(i)
        except KeyError:
            raise KeyError("Could not increment Sales Ping")

    def count_web_form(self, i=1):
        try:
            self.ping_stats["Web Form"].increment(i)
        except KeyError:
            raise KeyError("Could not increment Web Form")

    def count_open(self, i=1):
        try:
            self.total_stats["Total Open Inquiries"].increment(i)
        except KeyError:
            raise KeyError("Could not increment Total Open")

    def count_new_open(self, i=1):
        try:
            self.total_stats["New Open Inquiries"].increment(i)
        except KeyError:
            raise KeyError("Could not increment New Open")

    def count_new_closed(self, i=1):
        try:
            self.total_stats["New Closed Inquiries"].increment(i)
        except KeyError:
            raise KeyError("Could not increment New Closed")

    def count_existing_closed(self, i=1):
        try:
            self.total_stats["Existing Open Inquiries Closed"].increment(i)
        except KeyError:
            raise KeyError("Could not increment Existing Open Inquiries Closed")

    def calc_total_closed(self):
        count = self.total_stats["New Closed Inquiries"].get_count() + \
                self.total_stats["Existing Open Inquiries Closed"].get_count()
        self.total_stats["Total Closed Inquiries"].increment(count)
        return self.total_stats["Total Closed Inquiries"]

    def format_stats(self):
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
        change = self.stat_labels["Change Request"].get_count()
        access = self.stat_labels["Change Request - Access Level"].get_count()
        self.stat_labels["Change Request"].set_count(change + access)
        del self.stat_labels["Change Request - Access Level"]

        issue = self.stat_labels["Issue"].get_count()
        system = self.stat_labels["System Access Issue"].get_count()
        pdf = self.stat_labels["Issue/PDF"].get_count()

        issue_priority = self.stat_labels["Issue"].get_priority()
        self.stat_labels["Issues"] = Stat("Issues", issue + system + pdf, issue_priority)
        del self.stat_labels["Issue"]
        del self.stat_labels["System Access Issue"]
        del self.stat_labels["Issue/PDF"]

        citi1 = self.stat_labels["CITI Integration"].get_count()
        citi_priority = self.stat_labels["CITI Integration"].get_priority()
        citi2 = self.stat_labels["CITI Interface Errors"].get_count()
        del self.stat_labels["CITI Integration"]
        del self.stat_labels["CITI Interface Errors"]
        self.stat_labels["CITI"] = Stat("CITI", citi1 + citi2, citi_priority)
        assert "CITI" in self.stat_labels

        pings_in_stats = ["Welcome to Support", "Welcome Ping", "Sales"]
        i = self.ping_stats["Sales Pings"].get_priority() - len(pings_in_stats)
        for ping in pings_in_stats:
            self.ping_stats[ping] = self.stat_labels[ping]
            del self.stat_labels[ping]
            self.ping_stats[ping].set_priority(i)
            i += 1

        # Calculate total Pings and Non-Pings
        total_non_pings = 0
        for stat in self.stat_labels:
            total_non_pings += self.stat_labels[stat].get_count()

        self.stat_labels["Total Non-Pings"] = Stat("Total Non-Pings", total_non_pings, issue_priority + 0.1)

        total_pings = self.ping_stats["User Inquiries"].get_count() + self.ping_stats["Demo Requests"].get_count() + \
            self.ping_stats["Welcome to Support"].get_count() + self.ping_stats["Welcome Ping"].get_count() + \
            self.ping_stats["Sales"].get_count() + self.ping_stats["New Organizations"].get_count() + \
            self.ping_stats["Sales Pings"].get_count() + self.ping_stats["Voicemails"].get_count() + \
            self.ping_stats["Support Pings"].get_count()

        # Combine pings
        sp = self.ping_stats["Sales Pings"].get_count()
        web = self.ping_stats["Web Form"].get_count()
        self.ping_stats["Sales Pings"].set_count(str(sp) + "(" + str(web) + ")")
        del self.ping_stats["Web Form"]

        self.ping_stats["Total Pings"].set_count(total_pings)

        # Calculate total open and closed
        self.total_stats["Overall Total New Inquiries"].set_count(total_pings + total_non_pings)
        self.calc_total_closed()
        self.total_stats["Total Open Inquiries"].increment(self.total_stats['New Open Inquiries'].get_count())
        # Get call info

    def count_stats(self, threads, members):
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
                filename = 'tools/logs/thread_lookup.csv'
                if config.TEST:
                    filename = 'tools/logs/test_thread_lookup.csv'
                out = open(filename, 'wb')
                writer = csv.writer(out)
                writer.writerow(["Date", "Subject", "Stats", "Members", "Counted Stat", "Closed?"])
            except IOError:
                util.print_error("Error. Could not open thread_lookup. Results will not be logged.")

        open_inquiries = {}
        for thread in threads:
            trd = threads[thread]
            if trd.is_good():
                if trd.is_admin_vm():
                    self.count_admin_vm()
                    if out is not None:
                        _write_stat_row(writer, trd, "Admin VM")
                elif trd.is_res_vm():
                    self.count_res_vm(trd.get_count())
                    if out is not None:
                        _write_stat_row(writer, trd, "Researcher VM")
                elif trd.is_finance_vm():
                    self. count_finance_vm()
                    if out is not None:
                        _write_stat_row(writer, trd, "Finance VM")
                elif trd.is_sales_vm():
                    self.count_sales_vm()
                    if out is not None:
                        _write_stat_row(writer, trd, "Sales VM")
                #  Handle Pings
                if trd.is_inquiry():
                    if trd.is_sales_ping():
                        self.count_web_form()
                        self.count_sales_ping()
                        if out is not None:
                            _write_stat_row(writer, trd, "Sales Pings")
                    else:
                        self.count_user_inquiry()
                        if out is not None:
                            _write_stat_row(writer, trd, "Inquiry")
                elif trd.is_demo():
                    if trd.is_sales_ping():
                        self.count_web_form()
                        self.count_sales_ping()
                        if out is not None:
                            _write_stat_row(writer, trd, "Sales Pings")
                    else:
                        if out is not None:
                            _write_stat_row(writer, trd, "Demo")
                        self.count_demo()
                elif trd.is_support_ping():
                    if out is not None:
                        _write_stat_row(writer, trd, "Support Ping")
                    self.count_support_ping()
                elif trd.is_sales_ping():
                    self.count_sales_ping()
                    if out is not None:
                        _write_stat_row(writer, trd, "Sales Pings")
                elif trd.is_new_org():
                    self.count_new_org()
                    if out is not None:
                        _write_stat_row(writer, trd, "New Org")

                # Handle Non-pings, Count lowest priority stat only
                sorted_stats = _sort_stats(trd.get_stats(), self.stat_labels)
                if len(sorted_stats) > 0:
                    counted_stat = sorted_stats[0]
                    self.stat_labels[counted_stat].increment()
                    # Sales inquiries are considered Pings
                    if counted_stat not in ["Sales"]:
                        if trd.is_closed():
                            self.count_new_closed()
                        else:
                            self.count_new_open()
                            open_inquiries[trd.get_id()] = mail.OpenInquiry(trd.get_id(), trd.get_subject())
                    if out is not None:
                        _write_stat_row(writer, trd, counted_stat)

                    for mem in trd.get_members():
                        for stat in sorted_stats:
                            members[mem].increment_stat(self.stat_labels[stat].get_priority())
                else:
                    # TODO Error Log
                    pass

        if out is not None:
            out.close()
        return open_inquiries

    def get_support_calls(self):
        """
        Attempts to obtain the number of support calls from the Enrollment Dashboard. Updates the CALL_STATS with
        correct values.
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

        self.call_stats["Sessions"].set_count(sessions)
        self.call_stats["Sales Calls"].set_count(sales_calls)
        self.call_stats["Demos"].set_count(sales_demos)
        self.call_stats["Institutions"].set_count(demo_institutions)

    def draft_message(self, start_date, end_date):
        """
        Writes an email summarizing weekly statistics.
        :param start_date: Earliest date for which stats were counted.
        :param end_date: Latest date (inclusive) for which stats were counted.
        :return: the email text
        """
        initials = raw_input("Enter your initials...  ")

        end = end_date.strftime('%m/%d/%Y')
        end_weekday = end_date.strftime('%a')
        start = start_date.strftime('%m/%d/%Y')
        start_weekday = start_date.strftime('%a')

        total_pings = self.ping_stats["Total Pings"].get_count()
        total_non_pings = self.stat_labels["Total Non-Pings"].get_count()
        sp = self.ping_stats["Sales Pings"].get_count()
        pings_less_sales = str(total_pings - int(sp[:str(sp).find("(")]))  # This removes web forms.
        total = str(total_pings + total_non_pings)

        txt = "Andy, \n\n" \
            "Below are the requested statistics from " + str(start_weekday) + ", " + str(start) + " to " + \
              str(end_weekday) + ", " + str(end) + ":\n\n" \
            "Pings (includes New Orgs; no Sales Pings): " + pings_less_sales + "\n" \
            "Non-Pings: " + str(total_non_pings) + "\n" \
            "Sales Inquiries: " + str(self.ping_stats["Sales Pings"].get_count()) + "\n" \
            "Overall Total New Inquiries: " + total + "\n\n" \
            "Total New Inquiries (Non-Pings) Currently Open: " + \
              str(self.total_stats["New Open Inquiries"].get_count()) + "\n" \
            "Total New Inquiries (Non-Pings) Closed: " + \
              str(self.total_stats["New Closed Inquiries"].get_count()) + "\n" \
            "Total Existing Open Inquiries Closed: " + \
              str(self.total_stats["Existing Open Inquiries Closed"].get_count()) + \
              "\nTotal Open Inquiries: " + str(self.total_stats["Total Open Inquiries"].get_count()) + "\n" \
            "Total Closed Inquiries: " + str(self.total_stats["Total Closed Inquiries"].get_count()) + "\n\n" \
            "Total Researcher VMs: " + str(self.vm_stats[config.VM_RESEARCHER].get_count()) + "\n" \
            "Total Admin VMs: " + str(self.vm_stats[config.VM_ADMIN].get_count()) + "\n" \
            "Total Finance VMs: " + str(self.vm_stats[config.VM_FINANCE].get_count()) + "\n" \
            "Total Sales VMs: " + str(self.vm_stats[config.VM_SALES].get_count()) + "\n\n" \
            "Total # of Sessions: " + str(self.call_stats["Sessions"].get_count()) + "\n" \
            "Total # of Sales Calls: " + str(self.call_stats["Sales Calls"].get_count()) + "\n" \
            "Total # of Demo Calls: " + str(self.call_stats["Demos"].get_count()) + " " + \
              str(self.call_stats["Institutions"].get_count()) + "\n" \
            "Total # of Retention Calls: " + str(get_retention_calls()) + "\n\n" \
            "Cumulative Weekly Statistics can be accessed via the following sheet:\n" \
            "https://docs.google.com/a/irbnet.org/spreadsheets/d/" + config.WEEKLY_STATS_SPREADSHEET_ID + "\n\n" \
            "Retention information can be accessed via the following sheet\n" \
            "https://docs.google.com/spreadsheets/d/" + config.RETENTION_SPREADSHEET_ID + "\n\n" \
            "Let me know if you have any questions. Thanks!\n\n" + initials

        return txt

    def update_weekly_support_stats(self, service, spreadsheet_id, sheet_id=0):
        """
        Updates the specified sheet with aggregated statistics info by inserting a new column between columns B and C.
        The new column has the format.
            Date

            STAT_LABELS


            PING_STATS


            TOTAL_STATS

            CALL_STATS

            VM_STATS

        :param service: Authorized Google Sheets service to access the Sheets API
        :param spreadsheet_id: Spreadsheet ID for the sheet that will be updated.
        :param sheet_id: Google Sheet ID for the sheet that will be updated.
        :return: None
        """

        values = [(util.serial_date(datetime.datetime.today()), 'DATE')]
        values.extend(add_stats(self.stat_labels))
        values.extend([("", 'STRING'), ("", 'STRING')])
        values.extend(add_stats(self.ping_stats))
        values.extend([("", 'STRING'), ("", 'STRING')])
        values.extend(add_stats(self.total_stats))
        values.append(("", 'STRING'))
        values.extend(add_stats(self.call_stats))
        values.append(("", 'STRING'))
        values.extend(add_stats(self.vm_stats))
        insert_column_request = googleAPI.insert_column_request(sheet_id, values, 0, len(values), 2, 3)
        try:
            googleAPI.spreadsheet_batch_update(service, spreadsheet_id, insert_column_request)
        except HttpError:
            util.print_error("Failed to update Weekly Support Stats. ")
            try:
                out_file = open('stats.csv', 'wb')
                out = csv.writer(out_file)
                for (v, t) in values:
                    print (v, t)
                    out.writerow([v])
                print "Use stats.csv to update Weekly Support Stats Sheet"
                out_file.close()
            except IOError:
                util.print_error("Failed to write stats.csv. See stats below.")
                for (v, t) in values:
                    print v

    def get_totals(self):
        total_pings = self.ping_stats["Total Pings"].get_count()
        total_non_pings = self.stat_labels["Total Non-Pings"].get_count()
        sales_pings = self.ping_stats["Sales Pings"].get_count()
        pings_less_sales = total_pings - int(sales_pings[:str(sales_pings).find("(")])  # This removes web forms.
        total = total_pings + total_non_pings

        return total, total_non_pings, pings_less_sales

    @staticmethod
    def draft_html_message(support_counter, gov_counter, start_date, end_date):
        """
        Writes an email summarizing weekly statistics.
        :param gov_counter: gov StatCounter
        :param support_counter: support StatCounter
        :param start_date: Earliest date for which stats were counted.
        :param end_date: Latest date (inclusive) for which stats were counted.
        :return: the email text
        """
        initials = raw_input("Enter your initials...  ")

        total, total_non_pings, pings_less_sales = support_counter.get_totals()
        gov_total, gov_total_non_pings, gov_pings_less_sales = gov_counter.get_totals()

        end = end_date.strftime('%m/%d/%Y')
        end_weekday = end_date.strftime('%a')
        start = start_date.strftime('%m/%d/%Y')
        start_weekday = start_date.strftime('%a')
        support_pings_less_sales = pings_less_sales
        gov_pings_less_sales = gov_pings_less_sales
        support_total_non_pings = total_non_pings
        gov_total_non_pings = gov_total_non_pings
        support_sales_pings = support_counter.ping_stats["Sales Pings"].get_count()
        gov_sales_pings = gov_counter.ping_stats["Sales Pings"].get_count()
        support_total = total
        gov_total = gov_total

        support_new_open = support_counter.total_stats["New Open Inquiries"].get_count()
        gov_new_open = gov_counter.total_stats["New Open Inquiries"].get_count()
        support_new_closed = support_counter.total_stats["New Closed Inquiries"].get_count()
        gov_new_closed = gov_counter.total_stats["New Closed Inquiries"].get_count()
        support_existing_open_closed = support_counter.total_stats["Existing Open Inquiries Closed"].get_count()
        gov_existing_open_closed = gov_counter.total_stats["Existing Open Inquiries Closed"].get_count()
        support_total_open = support_counter.total_stats["Total Open Inquiries"].get_count()
        gov_total_open = gov_counter.total_stats["Total Open Inquiries"].get_count()
        support_total_closed = support_counter.total_stats["Total Closed Inquiries"].get_count()
        gov_total_closed = gov_counter.total_stats["Total Closed Inquiries"].get_count()

        support_vm_researcher = support_counter.vm_stats[config.VM_RESEARCHER].get_count()
        gov_vm_researcher = gov_counter.vm_stats[config.VM_RESEARCHER].get_count()
        support_vm_admin = support_counter.vm_stats[config.VM_ADMIN].get_count()
        gov_vm_admin = gov_counter.vm_stats[config.VM_ADMIN].get_count()
        support_vm_finance = support_counter.vm_stats[config.VM_FINANCE].get_count()
        support_vm_sales = support_counter.vm_stats[config.VM_SALES].get_count()
        sessions = support_counter.call_stats["Sessions"].get_count()

        sales_calls = support_counter.call_stats["Sales Calls"].get_count()
        demos = support_counter.call_stats["Demos"].get_count() + " " + support_counter.call_stats[
            "Institutions"].get_count()
        retention_calls = get_retention_calls()

        weekly_stats_sheet_id = config.WEEKLY_STATS_SPREADSHEET_ID
        retention_sheet_id = config.RETENTION_SPREADSHEET_ID

        html = """
        <p>Andy,</p> 
            <p>
                Below are the requested statistics from {start_weekday}, {start} to {end_weekday}, {end}
            </p>
            
                <table width="50%">
                <thead>
                    <tr>
                        <td width="60%"></td>
                        <td width="15%"><strong>Support</strong></td>
                        <td width="25%"><strong>Gov Support</strong></td>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td>Pings (includes New Orgs; no Sales Pings):</td>
                        <td>{support_pings_less_sales}</td>
                        <td>{gov_pings_less_sales}</td>
                    </tr>
                    <tr>
                        <td>Non-Pings:</td>
                        <td>{support_total_non_pings}</td>
                        <td>{gov_total_non_pings}</td>
                    </tr>
                    <tr>
                        <td>Sales Inquiries:</td>
                        <td>{support_sales_pings}</td>
                        <td>{gov_sales_pings}</td>
                    </tr>
                    <tr>
                        <td>Overall Total New Inquiries:</td>
                        <td>{support_total}</td>
                        <td>{gov_total}</td>
                    </tr>
                    <tr>
                        <td colspan="2">&nbsp;</td>
                    </tr>
                        <td>Total New Inquiries (Non-Pings) Currently Open:</td>
                        <td>{support_new_open}</td>
                        <td>{gov_new_open}</td>
                    </tr>
                    <tr>
                        <td>Total New Inquiries (Non-Pings) Closed:</td>
                        <td>{support_new_closed}</td>
                        <td>{gov_new_closed}</td>
                    </tr>
                    <tr>
                        <td>Total Existing Open Inquiries Closed:</td>
                        <td>{support_existing_open_closed}</td>
                        <td>{gov_existing_open_closed}</td>
                    </tr>
                    <tr>
                        <td>Total Open Inquiries:</td>
                        <td>{support_total_open}</td>
                        <td>{gov_total_open}</td>
                    </tr>
                    <tr>
                        <td>Total Closed Inquiries:</td>
                        <td>{support_total_closed}</td>
                        <td>{gov_total_closed}</td>
                    </tr>
                    <tr>
                        <td colspan="2">&nbsp;</td>
                    </tr>
                    <tr>
                    <td>Total Researcher VMs:</td>
                        <td>{support_vm_researcher}</td>
                        <td>{gov_vm_researcher}</td>
                    </tr>
                    <tr>
                        <td>Total Admin VMs:</td>
                        <td>{support_vm_admin}</td>
                        <td>{gov_vm_admin}</td>
                    </tr>
                    <tr>
                        <td>Total Finance VMs:</td>
                        <td colspan="2">{support_vm_finance}</td>
                    </tr>
                    <tr>
                        <td>Total Sales VMs:</td>
                        <td colspan="2">{support_vm_sales}</td>
                    </tr>
                    <tr>
                        <td colspan="2">&nbsp;</td>
                    </tr>
                    <tr>
                        <td>Total # of Sessions:</td>
                        <td colspan="2">{sessions}</td>
                    </tr>
                    <tr>
                        <td>Total # of Sales Calls:</td>
                        <td colspan="2">{sales_calls}</td>
                    </tr>
                    <tr>
                        <td>Total # of Demo Calls:</td>
                        <td colspan="2">{demos}</td>
                    </tr>
                    <tr>
                        <td>Total # of Retention Calls:</td>
                        <td colspan="2">{retention_calls}</td>
                    </tr>
                    </tbody>
            </table>
            
            <p>
               Cumulative Weekly Statistics can be accessed via the following sheet:<br />
               https://docs.google.com/a/irbnet.org/spreadsheets/d/{weekly_stats_sheet_id}
            </p>
            <p>
                Retention information can be accessed via the following sheet<br />
                https://docs.google.com/spreadsheets/d/{retention_sheet_id}
            </p>
            
            <p>
                Let me know if you have any questions. Thanks!
            </p>
            
            <p>{initials}</p>
        """.format(**locals())

        return MIMEText(html, 'html')
