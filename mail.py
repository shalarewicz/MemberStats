# Need to install the date util library
# pip install python-dateutil TODO Does this need to be done for everyone? Include in README
from dateutil import parser
from members import MEMBERS
from stats import STAT_LABELS
from util import is_mbox
import mailbox


class Message(object):
    """Represents a single Gmail message"""
    def __init__(self, message):
        """
        Creates a new message object from the provided GMail message. A Gmail message can be obtained via an
        mbox export or querying the inbox directly using the Google Mail API. The created message contains the
        following fields:
            Thread ID - GMail Thread ID or "" TODO Use None instead?
            To - To address or ""
            From - From address or ""
            Subject - Subject string or ""
            Labels - List of labels applied to the message
            Date - datetime object or None TODO Include a check when using date. Create log for None dates encountered
        :param message: GMail message
        """
        self.thread_id = message['X-GM-THRID']
        if self.thread_id is None:
            self.thread_id = ""
            # todo Should we throw an error here or just log it?

        self.to = message['To']
        if self.to is None:
            self.to = ""

        # Whitespace is a result of long strings in mbox file
        subject = message['Subject']
        if subject is None:
            self.subject = ""
        else:
            self.subject = subject.replace('\n', '').replace('\r','')

        from_address = message['From']
        if from_address is None:
            self.from_address = ""
        elif '<' not in from_address or '>' not in from_address:
            self.from_address = from_address.lower()
        else:
            # Extract email address from string containing "<email>"
            self.from_address = from_address[from_address.find("<") + 1, from_address.find(">")].lower()

        self.counts = not (self.is_internal() or self.is_from_support())

        labels = message['X-Gmail-Labels'].replace('\n', '').replace('\r', '')
        if labels is None:
            self.labels = []
        else:
            self.labels = labels.split(",")

        date = message['Date']
        if date is not None:
            self.date = parser.parse(date)

    def _is_spam(self):
        """
        Checks to see if the messages from address contains an address or substring indicative
        of a 'spam' message that can be ignored for stats. The following items are currently
        checked '<MAILER-DAEMON@LNAPL005.HPHC.org>' 'Mail Delivery System',
        'dmrn_exceptions@dmrn.dhhq.health.mil', and '<supportdesk@irbnet.org>'
        :return: true if the address matches any know spam strings or if the address is the empty string
        """
        # TODO Move outside of class
        spam = ["<MAILER-DAEMON@LNAPL005.HPHC.org>", "Mail Delivery System",
                "dmrn_exceptions@dmrn.dhhq.health.mil", "<supportdesk@irbnet.org>", ""]
        return any(s in self.from_address for s in spam)

    def _is_idea(self):
        """
        :return: True if the messages to address is to "ideas@irbet.org"
        """
        # TODO Move outside of class
        return "<ideas@irbnet.org>" in self.to

    def is_internal(self, address=from_address):
        """
        :return: True if the Message's from address matches an IRBNet personal email address
        # TODO Why do we rule out all of these addresses??? What happens if we don't?
        """
        # TODO Move outside of class
        internal = ["support@irbnet.org", "ideas@irbnet.org", "noreply@irbnet.org",
                    "supportdesk@irbnet.org", "techsupport@irbnet.org", "report_heartbeat@irbnet.org",
                    "report_monitor@irbnet.org", "alerts@irbnet.org", "wizards@irbnet.org",
                    "reportmonitor2@irbnet.org", ]
        return "irbnet.org" in address and all(x not in address for x in internal)

    def is_to_from_support(self):
        # TODO Move outside of class
        return "support@irbnet.org" in self.to and "support@irbnet.org" in self.from_address

    def is_from_support(self):
        return "support@irbnet.org" in self.from_address

    def extract_labels(self):
        stats = set()
        members = set()
        for label in self.labels:
            if label in STAT_LABELS:
                stats.add(label)
            if label in MEMBERS:
                members.add(label)
        return stats, members

    def get_thread_id(self):  # TODO Getters allow for easier changes in future and prevents rep exposure
        return self.thread_id

    def get_labels(self):
        return self.labels

    def get_date(self):
        return self.date

    def get_from_address(self):
        return self.from_address

    def get_to(self):
        return self.to

    def get_subject(self):
        return self.subject

    def counts(self):
        return self.counts


class Thread(object):

    closed_labels = ["Waiting on", "TO DO", "To Call"]

    """docstring for Thread"""
    def __init__(self, message):
        super(Thread, self).__init__()  # TODO necessary?

        self.id = message.get_thread_id()

        self.stat_labels, self.member_labels = message.extract_labels()
        self.good_thread = True
        self.last_contact_date = None
        self.oldest_date = None
        if message.counts():
            self.last_contact_date = message.get_date()
            self.oldest_date = message.get_date()
        self.check_in_date = None
        self.check_in = False  # TODO Remove and just ask if check_in_date is not None?
        self.non_ping = len(self.stat_labels) > 0
        self.demo = False
        self.inquiry = False
        self.vm = False
        self.new_org = False
        self.sales_ping = False
        self.message_count = 1
        self.checked = False
        self.closed = True
        self._evaluate(message)

    def _evaluate(self, message):
        subject = message.get_subject()
        if not self.non_ping and self.message_count == 2:
            if "IRBNet Demo Request" in subject:
                self.demo = True
                self.good_thread = True
            if "IRBNet Inquiry From" in subject:
                self.inquiry = True
                self.good_thread = True
            if "IRBNet Help Desk Inquiry" in subject:
                if "noreply@irbnet.org" not in message.from_address():
                    self.message_count -= 1
                self.vm = True
                self.good_thread = True
            if "New Organizations" in message.get_labels():
                self.new_org = True
                self.good_thread = True
            if "Sales Ping" in message.get_labels:
                self.sales_ping = True
                self.good_thread = True
                self.checked = True
        elif not self.checked:
            if message.is_to_from_support() and "New Organizations" not in self.stat_labels:
                self._should_it_count(message, "to and from Support")
            elif (message.is_internal() or (message.from_support() and message.is_internal(message.get_to)
                                            and "Sales Pings" not in self.stat_labels)):
                self._should_it_count(message, "Internal")

    def _should_it_count(self, message, type):
        # TODO Move this out of the thread class and just return a list of threads that need to
        #  be checked
        if not COUNT_ALL or COUNT_NONE:
            self.checked = True
            print "\nFound ", type, " email. Should the following message be counted?\n",\
                "\nFrom: " + message.get_from_address(),\
                "\nTo: " + message.get_to(),\
                "\nSubject: " + message.get_subject(),\
                "\nDate: " + message.get_oldest_date(),\
                "\nLabels:" + message.get_date()
            answer = raw_input("Y/N?    ").lower()
            if answer == "y":
                print "Thread will be counted."
                self.good_thread = True
            elif answer == "n":
                print "Thread won't be counted."
                self.good_thread = False
            else:
                print "Answer not recognized."
                self._should_it_count(message, type)
        elif COUNT_NONE:
            self.good_thread = False
        else:
            self.good_thread = True

    def add_message(self, message):
        """

        :param message:
        :return:
        """
        self.message_count += 1

        new_stats, new_members = message.extract_labels()
        for label in new_stats:
            self.stat_labels.add(label)

        for label in new_members:
            self.member_labels.add(label)

        new_date = message.date
        if new_date < self.oldest_date:
            self.oldest_date = new_date

        # If the message is from a member update the last contact date
        if message.counts() and (new_date > self.last_contact_date):
            self.last_contact_date = new_date

        if "check-in call" in message.get_labels():
            self.check_in = True
            if self.check_in_date < new_date:
                self.check_in_date = new_date

        # Mark the thread as open if it meets the criteria
        if self.closed and any(l in Thread.closed_labels for l in message.get_labels()):
            self.closed = False

        self.non_ping = len(self.stat_labels) > 0

        self._evaluate(message)

    def dont_count(self):
        self.good_thread = False

    def message_type(self):
        if self.demo:
            return "Demo"
        elif self.inquiry:
            return "Inquiry"
        elif self.vm:
            return "Voicemail"
        elif self.non_ping:
            return "Non Ping: " + self.stat_labels
        elif self.check_in:
            return "Check In"
        else:
            return "Unknown"

    def get_oldest_date(self):
        return self.oldest_date

    def get_count(self):
        return self.message_count

    def get_id(self):
        return self.id

    def get_stats(self):
        return self.stat_labels

    def get_members(self):
        return self.member_labels

    def is_good(self):
        return self.good_thread

    def is_non_ping(self):
        return self.non_ping

    def is_inquiry(self):
        return self.inquiry

    def is_demo(self):
        return self.demo

    def is_vm(self):
        return self.vm

    def is_closed(self):
        return self.closed

    def __str__(self):
        return "Thread ID: " + self.id + ", Type: " + self.message_type() + ", Count: " + str(self.message_count)


class OpenInquiry:
    """
    Represents and open and countable thread in the Support Inbox. Each
    thread is identified by the THREAD ID assigned by GMail.
    """
    def __init__(self, id, subject):
        self.id = id
        self.subject = subject

    def __repr__(self):
        return "< " + str(self.id) + ", " + self.subject + ">"

    def write(self):
        return str(self.id) + "\n" + self.subject

    def __eq__(self, other):
        return isinstance(other, OpenInquiry) and self.id == other.id

    @staticmethod
    def from_file(filename):
        """

        :param filename: File containing a list of open inquiries. Must be formatted properly
                         THREAD_ID
                         SUBJECT
        :return: set of all open inquires
        :raise: IOError if the file is not formatted properly or if the file cannot be read
        """
        input = open(filename, 'r')

        threads = {}
        while input:
            id = input.readline().strip()
            subject = input.readline().strip()
            threads[id] = OpenInquiry(id, subject)

        input.close()
        return threads

    @staticmethod
    def _from_mbox(filename):
        inbox = {}
        if is_mbox(filename):
            for message in mailbox.mbox(filename):
                id = message['X-GM-THRID']
                subject = message['Subject']

                if not any([id, subject] is None):
                    inbox[id] = subject
        else:
            raise IOError(filename + " is not a valid mbox file")

        return inbox

    @staticmethod
    def update(open_inquiries, filename):
        num_open = 0
        num_closed = 0
        to_delete = []
        current = OpenInquiry._from_mbox(filename)
        for thread in open_inquiries:
            if thread.id in current:
                num_open += 1
            else:
                num_closed += 1
                to_delete.append(thread)

        for thread in to_delete:
            del open_inquiries[thread]
            #TODO Check to make sure we don't have to return the updated dictionary

        return num_open
        return num_closed
        STAT_LABELS["Total Open Inquires"] = num_open
        STAT_LABELS["Existing Open Inquiries Closed"] = num_closed



