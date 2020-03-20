from googleAPI import get_range
import util


MEMBERS = {}  # Member dictionary {short name: Member}


class Member(object):
    """
    A Member allows for the tracking tracking member specific statistics, check in date and last contact date.

    Attributes:
        name: str
            member's short name
        last_contact: datetime:
            date the member last made contact with Support
        check_in: datetime
            date Support last conducted a check-in call
        stats: lst(int)
            Existing statistics from read from the Member Stats tab.
    """

    def __init__(self, name, last_contact, check_in, stats):
        """
        Creates a new Member.
        :param name: member's short name
        :param last_contact: date the member last made contact with Support
        :param check_in: date Support last conducted a check-in call
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
        """
        Compares the provided dates against the members record in MEMBERS and updates if necessary
        :param member:
        :param last_contact:
        :param check_in:
        :return: None
        """
        try:
            MEMBERS[member].update_last_contact(last_contact)
            MEMBERS[member].update_check_in(check_in)
        except KeyError:
            util.print_error("Error: Failed to updates dates for member (Member nof found): " + member.name)

    def get_stats(self):
        return self.stats

    def increment_stat(self, index):
        """
        Increments members.stats[index] by one
        :param index: index of the stat to be updated.
        :return: None
        """
        try:
            self.stats[index] += 1
        except IndexError:
            util.print_error('Error: IndexError during counting for ' + self.name + ': ' + str(index))
            util.print_error(self.get_stats())
            # todo log

    @staticmethod
    def create_stat_row(mem):
        """
        Creates a row which can be written to the Member Stats tab on the Retention sheet.
        Response format = [name, last_contact, check_in, stats]
        :param mem: str Member for which stat row would be created.
        :return: lst: a list with the response format listed above
        """
        if mem.last_contact is None:
            last_contact = ("", 'STRING')
        else:
            last_contact = (util.serial_date(mem.last_contact), 'DATE')

        if mem.check_in is None:
            check_in = ("", 'STRING')
        else:
            check_in = (util.serial_date(mem.check_in), 'DATE')

        result = [(mem.name, 'STRING'), last_contact, check_in]

        for stat in mem.stats:
            result.append((stat, 'NUMBER'))

        return result

    @staticmethod
    def read_members(rng, sheet_id, sheet_api, header_index, short_name_range):
        """
        Reads in member data from the specified range and google sheet and returns a dictionary of members and statistic
        labels that will be counted
        :param rng: Range to be read from sheet_id
        :param sheet_id: Google sheet_id for the sheet to be read
        :param sheet_api: Google sheets API used to read the sheet
        :param header_index: Start index (inclusive) of values that will be returned from header
        :param short_name_range: Named Range in Retention sheet for short names. If mem_stats_sheet does not contain a
                                 name in this range it will be added to the mem_stats_sheet
        :return: member dictionary with (k,v) = (name, Member() object), header[header_index:] from rng
        """
        data = get_range(rng, sheet_id, sheet_api)

        stat_header = data[0][header_index:]

        for member in data[1:]:
            try:
                name = member[0]
                contact = util.parse_date(member[1])
                check_in = util.parse_date(member[2])
                mem_stats = member[3:]
                parsed_stats = []
                for s in mem_stats:
                    parsed_stats.append(int(s))
                MEMBERS[name] = Member(name, contact, check_in, parsed_stats)
            except IndexError:
                print "The following member does not have a complete set of data on the " + rng + " tab." \
                      "Please update the member and rerun the script\n" + member
                exit()

        # Add any members not in the sheet
        for member in get_range(short_name_range, sheet_id, sheet_api):
            if member[0] not in MEMBERS.keys():
                # If the member is listed on the Retention tab but not Member Stats tab add a row of blank info
                MEMBERS[member[0]] = Member(member[0], None, None, [0] * len(stat_header))

        return MEMBERS, stat_header


class Admin(object):
    """
    Represents an administrator which last contact and check in information should be tracked.
    Attributes:
        name: str
        org: str
        emails: lst(str)
        :param last_contact: datetime
            date the member last made contact with Support
        :param check_in: datetime
            date Support last conducted a check-in call
        id: int
            Unique id
    """
    _id = 0

    def __init__(self, name, org, last_contact, check_in, emails):
        """
        Constructs a new admin.
        :param name: str
        :param org: str
        :param last_contact: datetime
        :param check_in: datetime
        :param emails: lst(str)
        """
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
        :return: admin dictionary with (k,v) = (id, Admin() object)
        """
        admins = {}
        admin_emails = {}

        data = get_range(rng, sheet_id, sheet_api)

        for admin in data[1:]:
            try:
                emails = [email.lower() for email in admin[6:9]]
                name = admin[1]
                org = admin[0]
                check_in = util.parse_date(admin[4])
                last_contact = util.parse_date(admin[5])

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
        """
        Compares the provided dates against the admin and updates if necessary
        :param admin: Admin
        :param last_contact: datetime
        :param check_in: datetime
        :return: None
        """
        admin.update_last_contact(util.serial_date(last_contact))
        admin.update_check_in(util.serial_date(check_in))

    @staticmethod
    def create_stat_row(admin):
        """
        Creates a row which can be written to the Support Outreach Admins tab on the Retention sheet.
        Response format = [last_contact, check_in]
        Dates will have format mm/dd/yyyy
        :param admin: str Admin for which stat row would be created.
        :return: lst: a list with the response format listed above
        """
        if admin.last_contact is None:
            last_contact = ("", 'STRING')
        else:
            last_contact = (util.serial_date(admin.last_contact), 'DATE')

        if admin.check_in is None:
            check_in = ("", 'STRING')
        else:
            check_in = (util.serial_date(admin.check_in), 'DATE')

        result = [check_in, last_contact]
        return result
