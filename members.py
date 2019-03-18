from googleAPI import get_range
from dateutil import parser
from sys import exit
import stats

MEMBERS = {}


class Member(object):
    """
    A Member allows for the tracking tracking member specific statistics, check in date and last contact date.
    """

    def __init__(self, name, last_contact, check_in, stats):
        """
        Creates a new Member.
        :param name: member's short name  TODO Type Specifications?
        :param last_contact: date the member last made contact with Support
        :param check_in: date Support last conducted a checkin call
        :param stats: Existing statistics from the Member Stats tab.
        """
        self.name = name
        self.last_contact = last_contact
        self.check_in = check_in
        self.stats = stats

    def set_check_in(self, date):
        """
        Updates the admins check in date if date is more recent thant he existing date.
        :param date: potential new date
        """
        if date > self.check_in: self.check_in = date

    def set_last_contact(self, date):
        """
        Updates the admins last_contact date if date is more recent thant he existing date.
        :param date: potential new date
        """
        if date > self.last_contact: self.last_contact = date

    def set_stats(self, new_stats):
        """ TODO is this used?
        Overwrites existing member statistics with new statistics
        :param new_stats:
        """
        self.stats = new_stats

    def get_stats(self):
        return self.stats

    def increment_stat(self, index):
        self.stats[index] += 1 #TODO sTRING ERROR?????

    @staticmethod
    def read_members(range, sheet_id, sheet_api):
        """
        Reads in member data from the specified range and google sheet and returns a dictionary of members and statistic
        labels that will be counted
        :param range: Range to be read from sheet_id
        :param sheet_id: Google sheet_id for the sheet to be read
        :param sheet_api: Google sheets API used to read the sheet
        :return: member dictionary with (k,v) = (name, Admin() object), stat label dicationary (k,v) = (label, Stat)
        """
        members = {}
        data = get_range(range, sheet_id, sheet_api)

        stat_header = data[0][3:]
        stats.Stat.extract_labels(stat_header)

        for member in data[1:]:
            try:
                members[member[0]] = Member(member[0], parser.parse(member[1]), parser.parse(member[2]), member[3:])
            except IndexError:
                print "The following member does not have a complete set of data on the " + range + " tab." \
                      "Please update the member and rerun the script\n" + member
                exit()

        # Add any members not in the sheet
        for member in get_range('Short_Names', sheet_id, sheet_api):
            if member[0] not in members.keys():
                # If the member is listed on the Retention tab but not Member Stats tab and a row of blank info
                members[member[0]] = Member(member[0], None, None, [0] * len(stat_header))

        return members


class Admin(object):
    _id = 0
    """docstring for Admin TODO"""
    def __init__(self, name, org, last_contact, check_in, emails):  #TODO flip check in and last contact dates in sheet
        super(Admin, self).__init__() # TODO WHAT IS THIS?

        self.name = name
        self.org = org
        self.emails = emails
        self.check_in = check_in
        self.last_contact = last_contact
        self.id = Admin._id
        Admin._id += 1

    @staticmethod
    def read_admins(range, sheet_id, sheet_api):
        """
        Reads in admin data from the specified range and google sheet and returns a dictionary of admins
        :param range: Range to be read from sheet_id
        :param sheet_id: Google sheet_id for the sheet to be read
        :param sheet_api: Google sheets API used to read the sheet
        :return: admin dictionary with (k,v) = (name, Admin() object)
        """
        admins = {}
        data = get_range(range, sheet_id, sheet_api)

        for admin in data[1:]:
            try:
                emails = [email.lower() for email in admin[6:9]]
                admins[admin[1]] = Admin(admin[1], admin[0], parser.parse(admin[4]), parser.parse(admin[5]), emails)

            except IndexError:
                print "The following admin does not have a complete set of data on the " + range + " tab." \
                      "This is likely the result of a missing phone number. " \
                      "Please update the member and rerun the script\n" + admin
                exit()

        return admins

    def set_check_in(self, date):
        """
        Updates the admins check in date if date is more recent thant he existing date.
        :param date: potential new date
        """
        if date > self.check_in: self.check_in = date

    def set_last_contact(self, date):
        """
        Updates the admins last_contact date if date is more recent thant he existing date.
        :param date: potential new date
        """
        if date > self.last_contact: self.last_contact = date
