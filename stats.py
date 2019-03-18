PING_STATS = ["Sales Pings", "New Organizations", "User Inquries", "Demo Requests", "Voicemails", "Total Pings"]
STAT_TOTALS = ["Overall Total New Inquiries", "New Open Inquiries",
               "New Closed Inquiries", "Existing Open Inquiries Closed", "Total Open Inquiries",
               "Total Closed Inquiries"]
STAT_LABELS = {}


# TODO stat label class???
def add_labels():
    for label in PING_STATS:
        STAT_LABELS[label] = Stat(label)

    for label in STAT_TOTALS:
        STAT_LABELS[label] = Stat(label)


def format_stats():
    change = STAT_LABELS["Change Request"].get_count()
    access = STAT_LABELS["Change Request - Access Level"].get_count()
    STAT_LABELS["Change Request"].set_count(change + access)
    del STAT_LABELS["Change Request - Access Level"]

    issue = STAT_LABELS["Issues"].get_count()
    system = STAT_LABELS["System Access Issue"].get_count()
    pdf = STAT_LABELS["Issue/PDF"].get_count()

    issue_priority = STAT_LABELS["Issues"]._priority
    STAT_LABELS["Issues"] = Stat("Issues", issue + system + pdf, issue_priority)
    del STAT_LABELS["Issue"]
    del STAT_LABELS["System Access Issue"]
    del STAT_LABELS["Issue/PDF"]

    citi1 = STAT_LABELS["CITI Integration"].get_count()
    citi_priority = STAT_LABELS["CITI Integration"]._priority
    citi2 = STAT_LABELS["CITI Interface Errors"].get_count()
    del STAT_LABELS["CITI Integration"]
    del STAT_LABELS["CITI Interface Errors"]
    STAT_LABELS["CITI"] = Stat(citi1 + citi2, citi_priority)

    # TODO Left off here. Never counted new orgs or Sales Pings in count()




    issue_priority = STAT_LABELS["Issue"]._priority
    STAT_LABELS["Total Non-Pings"] = Stat("Total Non-Pings", 0, issue_priority + 0.1)
    STAT_LABELS["non ping space"] = Stat("", issue_priority + 0.2)
    STAT_LABELS["Category"] = Stat("Category", issue_priority + 0.3)

    STAT_LABELS["ping space 1"] = Stat("", STAT_LABELS["Total Pings"]._priority + 0.1)
    STAT_LABELS["ping space 2"] = Stat("", STAT_LABELS["Total Pings"]._priority + 0.2)
    STAT_LABELS["closed space"] = Stat("", STAT_LABELS["Total Closed Inquiries"]._priority + 0.3)


def extract_labels(data):
    for stat in data:
        STAT_LABELS[stat] = Stat(stat)

        STAT_LABELS["Sales Pings"] = Stat()  # TODO this used to have priority 30?

    return STAT_LABELS


def add_call_info(stats, sessions, sales_calls, demos):
    stats["Total # of Sessions"] = Stat(sessions)
    stats["Total # of Sales Calls"] = Stat(sales_calls)
    stats["Total # of Demo Calls"] = Stat(demos)


# 15. The call sortStats(str list stats) -> str list returns the given list of stats
#  	  sorted by the stats assigned priority. Sorted lo --> hi priority
def _sort_stats(stats):
    less = []
    equal = []
    greater = []
    if len(stats) > 1:
        pivot = STAT_LABELS[stats[0]].priority
        for stat in stats:
            if STAT_LABELS[stat].priority < pivot:
                less.append(stat)
            if STAT_LABELS[stat].priority == pivot:
                equal.append(stat)
            if STAT_LABELS[stat].priority > pivot:
                greater.append(stat)
        return _sort_stats(less) + equal + _sort_stats(greater)
    else:
        return stats


def count_stats(threads):
    new_closed = 0
    new_open = 0
    open_inquires = set()
    for thread in threads:
        if thread.is_good():
            if thread.is_non_ping():
                if thread.is_inquiry():
                    STAT_LABELS["User Inquiries"].increment()
                if thread.is_demo():
                    STAT_LABELS["Demo Requests"].increment()
                if thread.is_vm():
                    STAT_LABELS["Voicemails"].set_count(thread.get_count())
                # TODO Used to check stat_label length and report error if zero. non_ping now determined by
                #  the length of stat_labels so uneccessary
            else:
                sorted_stats = _sort_stats(thread.get_stats())
                if sorted_stats is not None:
                    counted_stat = sorted_stats[0]
                    STAT_LABELS[counted_stat].increment()
                    # Sales Pings and Sales are considered Pings
                    if counted_stat not in ["Sales Pings", "Sales"]:
                        if thread.is_closed():
                            new_closed += 1
                        else:
                            new_open += 1
                            open_inquires.add(thread.get_id(), "Subject")  # TODO need to access message subject
                    try:
                        for mem in thread.get_members():
                            for stat in sorted_stats:
                                mem.increment_stat(STAT_LABELS[counted_stat]._priority)
                    except IndexError:
                        print "Mem Specific index error"
                        print mem.get_stats()  # TODO I don't think we can reference mem
                        print sorted_stats
                        # sys.exit() # TODO Log or exit?

                else:
                    # TODO Log error
                    pass

    return new_open, new_closed




class Stat:
    _id = -1 # First stat will have priority 0. This makes indexing easier when counting.

    def __init__(self, name, count=0, priority=_id):
        self.stat = name
        if priority == Stat._id:
            Stat._id += 1
        self._count = count
        self._priority = priority

    def increment(self):
        self._count += 1

    def decrement(self):
        self._count -= 1

    def set_count(self, x):
        self._count = x

    def get_count(self):
        return self._count

    def __str__(self):
        return self.stat + ": (" + str(self._count) + ", " + str(self._priority) + ")"

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



