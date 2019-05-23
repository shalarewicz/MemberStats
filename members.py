from googleAPI import get_range
from sys import exit
# import stats
import util
MEMBERS = {}


try:
    from dateutil import parser
except ImportError, e:
    print "ERROR: dateutil not found. Install the dateutil package by typing " \
          "'pip install python-dateutil' into the command line"


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

    def update_check_in(self, date):
        """
        Updates the member's check in date if date is more recent thant he existing date.
        :param date: potential new date
        """
        if date is None:
            pass
        elif date is not None and (self.check_in is None or date > self.check_in):
            self.check_in = date

    def update_last_contact(self, date):
        """
        Updates the member's last_contact date if date is more recent thant he existing date.
        :param date: potential new date
        """
        if date is not None and (self.last_contact is None or date > self.last_contact):
            self.last_contact = date

    @staticmethod
    def update_member_dates(member, last_contact, check_in):
        MEMBERS[member].update_last_contact(last_contact)
        MEMBERS[member].update_check_in(check_in)

    def set_stats(self, new_stats):
        """ TODO is this used?
        Overwrites existing member statistics with new statistics
        :param new_stats:
        """
        self.stats = new_stats

    def get_stats(self):
        return self.stats

    def increment_stat(self, index):
        try:
            self.stats[index] += 1
        except IndexError:
            print "IndexError during counting for " + self.name + ": " + str(index)
            print self.get_stats()
            # todo log

    @staticmethod
    def create_stat_row(mem):
        if mem.last_contact is None:
            last_contact = ""
        else:
            last_contact = mem.last_contact.strftime('%m/%d/%Y')

        if mem.check_in is None:
            check_in = ""
        else:
            check_in = mem.check_in.strftime('%m/%d/%Y')

        result = [mem.name, last_contact, check_in]
        result.extend(mem.stats)
        return result

    @staticmethod
    def read_members(rng, sheet_id, sheet_api):
        """
        Reads in member data from the specified range and google sheet and returns a dictionary of members and statistic
        labels that will be counted
        :param rng: Range to be read from sheet_id
        :param sheet_id: Google sheet_id for the sheet to be read
        :param sheet_api: Google sheets API used to read the sheet
        :return: member dictionary with (k,v) = (name, Member() object)
        """
        data = get_range(rng, sheet_id, sheet_api)

        stat_header = data[0][3:]

        for member in data[1:]:
            try:
                name = member[0]
                contact = util.parse_date(member[1])
                check_in = util.parse_date(member[2])
                mem_stats = member[3:]
                parsed_stats = []
                for s in mem_stats:
                    parsed_stats.append(int(s[0]))
                MEMBERS[name] = Member(name, contact, check_in, parsed_stats)
            except IndexError:
                print "The following member does not have a complete set of data on the " + rng + " tab." \
                      "Please update the member and rerun the script\n" + member
                exit()

        # Add any members not in the sheet
        for member in get_range('Short_Names', sheet_id, sheet_api):
            if member[0] not in MEMBERS.keys():
                # If the member is listed on the Retention tab but not Member Stats tab add a row of blank info
                MEMBERS[member[0]] = Member(member[0], None, None, [0] * len(stat_header))

        return MEMBERS, stat_header


class Admin(object):
    _id = 0
    """docstring for Admin TODO"""
    def __init__(self, name, org, last_contact, check_in, emails):  # TODO flip check in and last contact dates in sheet

        self.name = name
        self.org = org
        self.emails = emails
        self.check_in = check_in
        self.last_contact = last_contact
        self.id = Admin._id
        Admin._id += 1

    @staticmethod
    def read_admins(rng, sheet_id, sheet_api):
        """
        Reads in admin data from the specified range and google sheet and returns a dictionary of admins
        :param rng: Range to be read from sheet_id
        :param sheet_id: Google sheet_id for the sheet to be read
        :param sheet_api: Google sheets API used to read the sheet
        :return: admin dictionary with (k,v) = (name, Admin() object)
        """
        admins = {}
        admin_emails = {}

        data = get_range(rng, sheet_id, sheet_api)

        for admin in data[1:]:
            try:
                emails = [email.lower() for email in admin[6:9]]
                name = admin[1]
                org = admin[0]
                last_contact = util.parse_date(admin[4])
                check_in = util.parse_date(admin[5])

                new = Admin(name, org, last_contact, check_in, emails)
                admins[new.id] = new
                for email in emails:
                    admin_emails[email] = new.id

            except IndexError:
                print "The following admin does not have a complete set of data on the " + rng + " tab." \
                      "This is likely the result of a missing phone number. " \
                      "Please update the member and rerun the script\n" + admin
                exit()

        return admins, admin_emails

    def update_check_in(self, date):
        """
        Updates the admin's check in date if date is more recent thant he existing date.
        :param date: potential new date
        """
        if date is None:
            pass
        elif date is not None and (self.check_in is None or date > self.check_in):
            self.check_in = date

    def update_last_contact(self, date):
        """
        Updates the admin's last_contact date if date is more recent thant he existing date.
        :param date: potential new date
        """
        if date is not None and (self.last_contact is None or date > self.last_contact):
            self.last_contact = date

    @staticmethod
    def update_admin_dates(admin, last_contact, check_in):
        admin.update_last_contact(last_contact)
        admin.update_check_in(check_in)

    @staticmethod
    def create_stat_row(admin):
        if admin.last_contact is None:
            last_contact = ""
        else:
            last_contact = admin.last_contact.strftime('%m/%d/%Y')

        if admin.check_in is None:
            check_in = ""
        else:
            check_in = admin.check_in.strftime('%m/%d/%Y')

        result = [last_contact, check_in]
        return result
