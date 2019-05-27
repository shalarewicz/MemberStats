import sys
import time
import csv

import config
from mail import OpenInquiry
import util

import googleAPI
from googleapiclient.errors import HttpError  # TODO This shouldn't need to be handled here.


class Stat:

    _id = 0  # First stat will have priority 0. This makes indexing easier when counting.

    def __init__(self, name, count=0, priority=-1):
        if priority == -1:
            priority = Stat._id
            Stat._id += 1

        self.stat = name
        self._count = count
        self._priority = priority

    def increment(self, i=1):
        self._count += i

    def decrement(self, i=1):
        self._count -= i

    def set_count(self, x):
        self._count = x

    def get_count(self):
        return self._count

    def get_priority(self):
        return self._priority

    def set_priority(self, p):
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
    TOTAL_STATS["Total Closed Inquiries"].set_count(count)
    return count


def calc_total_open():
    count = TOTAL_STATS["New Open Inquiries"].get_count() - TOTAL_STATS["Existing Open Inquiries Closed"].get_count()
    TOTAL_STATS["Total Open Inquiries"].set_count(count)
    return count


def format_stats():
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

    # Calculate total Pings and Non-Pings
    total_non_pings = STAT_LABELS["Product Enhancement"].get_count() + STAT_LABELS["WIRB"].get_count() + \
        STAT_LABELS["Funding Status"].get_count() + STAT_LABELS["Alerts"].get_count() + \
        STAT_LABELS["CITI"].get_count() + \
        STAT_LABELS["IRBNet Resources"].get_count() + STAT_LABELS["Reports"].get_count() + \
        STAT_LABELS["Smart Forms"].get_count() + STAT_LABELS["Stamps"].get_count() + \
        STAT_LABELS["Letters"].get_count() + STAT_LABELS["Information"].get_count() + \
        STAT_LABELS["Change Request"].get_count() + STAT_LABELS["Issues"].get_count()

    total_pings = PING_STATS["User Inquiries"].get_count() + PING_STATS["Demo Requests"].get_count() + \
        STAT_LABELS["Welcome to Support"].get_count() + STAT_LABELS["Welcome Ping"].get_count() + \
        STAT_LABELS["Sales"].get_count() + PING_STATS["New Organizations"].get_count() + \
        PING_STATS["Sales Pings"].get_count() + \
        PING_STATS["Voicemails"].get_count()

    STAT_LABELS["Total Non-Pings"] = Stat("Total Non-Pings", total_non_pings, issue_priority + 0.1)

    STAT_LABELS["non ping space"] = Stat("", 0, issue_priority + 0.2)
    STAT_LABELS["Category"] = Stat("Category", 0, issue_priority + 0.3)

    # Combine pings
    sp = PING_STATS["Sales Pings"].get_count()
    web = PING_STATS["Web Form"].get_count()
    PING_STATS["Sales Pings"].set_count(str(sp) + "(" + str(web) + ")")
    del PING_STATS["Web Form"]

    ping_priority = PING_STATS["Total Pings"].get_priority()
    PING_STATS["Total Pings"] = Stat("Total Pings", total_pings, ping_priority)

    PING_STATS["ping space 1"] = Stat("", 0, ping_priority + 1)  # TODO can remove this with 3 writes
    PING_STATS["ping space 2"] = Stat("", 0, ping_priority + 2)  # TODO Use generators for each dict

    # Calculate total open and closed
    TOTAL_STATS["Overall Total New Inquiries"].set_count(total_pings + total_non_pings)
    calc_total_closed()
    calc_total_open()

    TOTAL_STATS["closed space"] = Stat("", 0, TOTAL_STATS["Total Closed Inquiries"].get_priority() + 1)

    # Get call info
    get_support_calls()


def extract_labels(data):
    for stat in data:
        STAT_LABELS[stat] = Stat(stat)

    return STAT_LABELS


# 15. The call sortStats(str list stats) -> str list returns the given list of stats
#  	  sorted by the stats assigned priority. Sorted lo --> hi priority
def _sort_stats(stats, dictionary):
    # TODO Insertion sort is faster for small lists.
    less = []
    equal = []
    greater = []
    if len(stats) > 1:
        pivot = dictionary[stats[0]].get_priority()
        for stat in stats:
            if dictionary[stat].get_priority() < pivot:
                less.append(stat)
            if dictionary[stat].get_priority() == pivot:
                equal.append(stat)
            if dictionary[stat].get_priority() > pivot:
                greater.append(stat)
        return _sort_stats(less, dictionary) + equal + _sort_stats(greater, dictionary)
    else:
        return stats


def _write_stat_row(writer, thread, stat, closed):
    writer.writerow([thread.get_oldest_date(), thread.get_subject(), thread.get_stats(), thread.get_members(),
                    stat, closed])


def count_stats(threads, members):
    out = open('Test/thread_lookup.csv', 'wb')  # TODO Surround with try statement
    writer = csv.writer(out)
    writer.writerow(["Date", "Subject", "Stats", "Members", "Counted Stat", "Closed?"])

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
                        _write_stat_row(writer, trd, "Sales Pings", trd.is_closed())
                    else:
                        count_user_inquiry()
                        _write_stat_row(writer, trd, "Inquiry", trd.is_closed())
                elif trd.is_demo():
                    if trd.is_sales_ping():
                        count_web_form()
                        count_sales_ping()
                        _write_stat_row(writer, trd, "Sales Pings", trd.is_closed())
                    else:
                        _write_stat_row(writer, trd, "Demo", trd.is_closed())
                        count_demo()
                elif trd.is_sales_ping():
                        count_sales_ping()
                elif trd.is_vm():
                    _write_stat_row(writer, trd, "Voicemail", trd.is_closed())
                    count_voicemail(trd.get_count())
            else:
                # Handle Non-pings, Count lowest priority stat only
                sorted_stats = _sort_stats(trd.get_stats(), STAT_LABELS)
                if len(sorted_stats) > 0:
                    counted_stat = sorted_stats[0]
                    STAT_LABELS[counted_stat].increment()
                    # Sales Pings and Sales are considered Pings
                    if counted_stat not in ["Sales"]:  # TODO what is this check for?
                        if trd.is_closed():
                            count_new_closed()
                        else:
                            count_new_open()
                            open_inquiries[trd.get_id()] = OpenInquiry(trd.get_id(), trd.get_subject())
                            # open_inquiries.add(trd.get_id(), trd.get_subject())  # TODO need to access subject

                        _write_stat_row(writer, trd, counted_stat, trd.is_closed())

                    for mem in trd.get_members():
                        for stat in sorted_stats:
                            members[mem].increment_stat(STAT_LABELS[stat].get_priority())
                else:
                    # TODO Error Log
                    pass

    out.close()
    return open_inquiries


def get_retention_calls():
    try:
        retention_calls = googleAPI.SHEETS_API.spreadsheets().values().get(
            spreadsheetId=config.SPREADSHEET_ID, range='Weekly_Check_In')\
            .execute().get('values', [])
        return str(retention_calls[0][0])
    except:
        print "Could not read number of check in calls."
        retention_calls = raw_input(
            "Enter the number of check in calls this week (cell B2 on Chart Data tab of Retention sheet).   ")
        return retention_calls


def _enter_calls():
    sessions = raw_input("\n\nEnter the Total # of Sessions...  ")
    sales_calls = raw_input("Enter the Total # of Sales Calls...  ")
    sales_demos = raw_input("Enter the Total # of Sales Demos...  ")
    demo_institutions = raw_input("Enter Sales Demo institutions...  ")
    return sessions, sales_calls, sales_demos, demo_institutions


def get_support_calls():
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
    initials = raw_input("Enter your initials...  ")

    today = time.strftime('%m/%d/%Y')
    today_weekday = time.strftime('%a')
    start_date = cutoff.strftime('%m/%d/%Y')
    start_weekday = cutoff.strftime('%a')

    # TODO Fail fast if STAT_LABELS dictionary incomplete.
    total_pings = PING_STATS["Total Pings"].get_count()
    total_non_pings = STAT_LABELS["Total Non-Pings"].get_count()
    sp = PING_STATS["Sales Pings"].get_count()
    pings_less_sales = str(total_pings - int(sp[:sp.find("(")]))  # This removes web forms.
    total = str(total_pings + total_non_pings)

    email = "Andy, \n\n " \
            "Below are the requested statistics from " + str(start_weekday) + ", " + start_date + " to " + \
            str(today_weekday) + ", " + str(today) + ":\n\n" \
            "Pings (includes New Orgs; no Sales Pings): " + pings_less_sales + "\n" \
            "Non-Pings: " + str(total_non_pings) + "\n" \
            "Sales Inquiries: " + str(PING_STATS["Sales Pings"].get_count()) + "\n" \
            "Overall Total New Inquiries: " + total + "\n\n" \
            "Total New Inquiries (Non-Pings) Currently Open: " + str(TOTAL_STATS["New Open Inquiries"].get_count()) + "\n" \
            "Total New Inquiries (Non-Pings) Closed: " + str(TOTAL_STATS["New Closed Inquiries"].get_count()) + "\n" \
            "Total Existing Open Inquiries Closed: " + str(TOTAL_STATS["Existing Open Inquiries Closed"].get_count()) + "\n" \
            "Total Open Inquiries: " + str(TOTAL_STATS["Total Open Inquiries"].get_count()) + "\n" \
            "Total Closed Inquiries: " + str(TOTAL_STATS["Total Closed Inquiries"].get_count()) + "\n\n" \
            "Total # of Sessions: " + str(CALL_STATS["Sessions"].get_count()) + "\n" \
            "Total # of Sales Calls: " + str(CALL_STATS["Sales Calls"].get_count()) + "\n" \
            "Total # of Demo Calls: " + str(CALL_STATS["Demos"].get_count()) + " " + str(CALL_STATS["Institutions"].get_count()) + "\n" \
            "Total # of Retention Calls: " + str(get_retention_calls()) + "\n\n" \
            "Cumulative Weekly Statistics can be accessed via the following sheet:\n" \
            "https://docs.google.com/a/irbnet.org/spreadsheets/d/" + config.WEEKLY_STATS_SHEET_ID + "\n\n" \
            "Retention information can be accessed via the following sheet\n" \
            "https://docs.google.com/spreadsheets/d/" + config.SPREADSHEET_ID + "\n\n" \
            "Let me know if you have any questions. Thanks!\n\n" + initials

    return email


def add_stats(dictionary):
    values = []
    for item in _sort_stats(dictionary.keys(), dictionary):
        values.append(dictionary[item].get_count())
    return values


def update_weekly_support_stats(service, sheet_id):
    values = []
    values.extend(add_stats(STAT_LABELS))
    values.extend(add_stats(PING_STATS))
    values.extend(add_stats(TOTAL_STATS))
    values.extend(add_stats(CALL_STATS))
    try:
        googleAPI.add_column(values, service, sheet_id, 2, 3, 0, 39)
    except HttpError:
        util.print_error("Failed to update Weekly Support Stats. ")
        try:
            out = csv.writer(open('stats.csv', 'w'))
            with out:
                for value in values:
                    out.writeln(value)
            print "Use stats.csv to update Weekly Support Stats Sheet"
        except IOError:
            util.print_error("Failed to write stats.csv. See stats below.")
            for value in values:
                print value
