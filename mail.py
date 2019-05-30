from members import MEMBERS
import stats
import config
import util


class Message(object):
    """Represents a single GMail message
    Attributes:
        thread_id: int
            GMail thread ID
        to: str
            Message recipient
        from_address: str
            Message sender
        subject: str
            Subject line
        date: datetime
            Subject Date
        labels: lst(str)
            All GMail labels associated with the message.
        counts: boolean
            True if the message is not internal or is from support
    """
    def __init__(self, message):
        """
        :param message: GMail message used to build the Message object.
        """
        self.thread_id = message['X-GM-THRID']
        if self.thread_id is None:
            self.thread_id = ""
            # TODO Log?

        self.to = message['To']
        if self.to is None:
            self.to = ""

        subject = message['Subject']
        if subject is None:
            self.subject = ""
        else:
            # Remove whitespace which can appear in messages with a long list of labels.
            self.subject = subject.replace('\n', '').replace('\r', '')

        from_address = message['From']
        if from_address is None:
            self.from_address = ""
        elif '<' not in from_address or '>' not in from_address:
            self.from_address = from_address.lower()
        else:
            # Extract email address from string containing "<email>"
            self.from_address = from_address[from_address.find("<") + 1: from_address.find(">")].lower()

        self.counts = not (self.is_internal() or self.is_from_support())

        self.labels = message['X-Gmail-Labels']

        date = message['Date']
        if date is not None:
            self.date = util.parse_date(date)

    def is_spam(self):
        """
        Checks to see if the messages from address contains an address or substring indicative
        of a 'spam' message that can be ignored for stats. Spam addresses listed in config.SPAM_EMAILS
        :return: True if the address matches any know spam strings or if the address is the empty string
        """
        return any(s in self.from_address for s in config.SPAM_EMAILS)

    def is_idea(self):
        """
        :return: True if the messages to address is to config.IDEAS_EMAIL
        """
        return config.IDEAS_EMAIL in self.to

    def is_internal(self):
        """
        :return: True if the Message's from address matches contains 'irbnet.org' and is not listed
        in config.INTERNAL_EMAILS.
        """
        return "irbnet.org" in self.from_address and all(x not in self.from_address for x in config.INTERNAL_EMAILS)

    def is_to_from_support(self):
        """

        :return: True if message is both to and from config.SUPPORT_EMAIL
        """

        return config.SUPPORT_EMAIL in self.to and config.SUPPORT_EMAIL in self.from_address

    def is_from_support(self):
        """

        :return: True if message is from config.SUPPORT_EMAIL
        """
        return config.SUPPORT_EMAIL in self.from_address

    def extract_labels(self):
        """
        Extracts all labels from the message labels that are in either in stats.STAT_LABELS or members.MEMBERS
        :return: 2 lists fro stat and member labels respectively.
        """
        statistics = set()
        members = set()
        for label in self.labels:
            if label in stats.STAT_LABELS:
                statistics.add(label)
            if label in MEMBERS:
                members.add(label)
        return statistics, members

    def get_thread_id(self):
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


class Thread(object):

    """Mail thread tracking the thread type for stats as well as basic thread information. Labels are only evaluated
    when a message is first added to the thread.

    Attributes:
        id : int
            GMail thread id
        subject : str
            Subject line of the thread
        message_count : int
            Number of messages counted in this thread
        stat_labels : set(str)
            Stat labels contained in the tread. Stat label is determined by those specified in stats.STAT_LABELS
        member_labels : set(str)
            Member labels contained in the tread. Member label is determined by those specified in members.MEMBERS
        good_thread : boolean
            True if the thread should count. (default = True)
        oldest_date : datetime
            Oldest date associated with the thread
        last_contact_date : datetime
            Most recent date associated with a message where message.counts = True
        check_in_date : datetime
            Most recent date associated with a message where 'check-in call' is in the list of message labels
        non_ping : boolean
            True if stat_labels contains at least one label
        demo : boolean
            True if subject contains 'IRBNet Demo Request' and message_count >= 2
        inquiry : boolean
            True if subject contains 'IRBNet Inquiry From' and message_count >= 2
        vm : boolean
            True if subject contains 'IRBNet Help Desk Inquiry' and count >= 2
        new_org : boolean
            True if message labels contains label 'New Organizations' and count >= 2
        sales_ping : boolean
            True if message labels contains label 'Sales Pings' and count >= 2
        checked : boolean
            True if the message has been manually checked if it should count (default=False)
        closed : boolean
            False if any of the message labels contained a phrase contained in config.OPEN_LABELS (default=True)

    """
    def __init__(self, message):
        """
        Constructs a new thread from the provided message.
        :param message: Message
            Used to construct the thread. Thread attributes are adjusted based on message information.
        """

        self.id = message.get_thread_id()

        self.stat_labels, self.member_labels = message.extract_labels()
        self.good_thread = True
        self.last_contact_date = None
        self.oldest_date = message.get_date()
        if message.counts:  # only update for messages not from support or an internal address
            self.last_contact_date = message.get_date()
        self.check_in_date = None
        self.non_ping = len(self.stat_labels) > 0
        self.demo = False
        self.inquiry = False
        self.vm = False
        self.new_org = False
        self.sales_ping = False
        self.message_count = 1
        self.checked = False
        self.closed = True
        self.subject = message.get_subject()
        self._evaluate(message)

    def _evaluate(self, message):
        """
        Determines a thread type by evaluating a message. Thread can be set as a demo, inquiry, new org, sales ping
        or voicemail.
        If a message is determined to be internal or to and from support, the user is asked whether or not the thread
        should count.
        Sets the check-in date and marks a message as not closed if applicable
        :param message: Message
            message being evaluated.
        :return: None
        """
        subject = message.get_subject()
        if not self.non_ping:
            if self.message_count == 2:
                # A count >= 2 avoids spam pings and genuine new organizations created by a member of the team.
                if "IRBNet Demo Request" in subject:
                    self.demo = True
                    self.good_thread = True
                elif "IRBNet Inquiry From" in subject:
                    self.inquiry = True
                    self.good_thread = True
                elif "New Organizations" in message.get_labels():
                    self.new_org = True
                    self.good_thread = True
            if "IRBNet Help Desk Inquiry" in subject:
                if "noreply@irbnet.org" not in message.get_from_address():
                    self.message_count -= 1  # Eliminates admin/researcher replies when total vm count is determined.
                self.vm = True
                self.good_thread = True
            if "Sales Pings" in message.get_labels():
                self.sales_ping = True
                self.good_thread = True
                self.checked = True
        elif not self.checked:
            if message.is_to_from_support() and not self.new_org:
                self.should_it_count(message, "to and from Support")
            elif (message.is_internal() or
                  (message.is_from_support() and message.is_internal() and not self.sales_ping)):
                self.should_it_count(message, "Internal")

        if "check-in call" in message.get_labels():
            if self.check_in_date is None or self.check_in_date < message.get_date():
                self.check_in_date = message.get_date()

        for l in config.OPEN_LABELS:
            if any(l in label for label in message.get_labels()):
                self.closed = False

    def should_it_count(self, message, message_type, override=False):
        """
        Prints basic message information and asks the user if a thread should count.
        Acceptable user responses include 'Y' | 'y' | 'N' | 'n'
        If an unrecognized response is entered the user is asked again.
        Sets self.good_thread as True if user responds 'Y' | 'y'
        Sets self.good_thread as False if user responds 'N' | 'n'
        Sets self.checked as True after a response is entered.

        If config.COUNT_NONE = True or config.COUNT_ALL = True, user prompt is skipped and good_thread is
        set accordingly
        :param message: Message
            Message from which information is extracted.
        :param message_type: str 'Support' | 'Internal
            Message type printed to provide user with more information.
        :param override: boolean
            True if system setting for COUNT_ALL and COUNT_NONE should be ignored and thread should be checked.
        :return: None
        """
        if override or not (config.COUNT_ALL or config.COUNT_NONE):
            self.checked = True
            print "\nFound ", message_type, " email. Should the following message be counted?\n",\
                "\nFrom: " + message.get_from_address(),\
                "\nTo: " + message.get_to(),\
                "\nSubject: " + message.get_subject(),\
                "\nDate: " + str(message.get_date()),\
                "\nLabels:" + str(message.get_labels())
            answer = raw_input("Y/N?    ").lower().strip()
            if answer == "y":
                print "Thread will be counted."
                self.good_thread = True
            elif answer == "n":
                print "Thread won't be counted."
                self.good_thread = False
            else:
                print "Answer not recognized."
                self.should_it_count(message, message_type)
        elif config.COUNT_NONE:
            self.good_thread = False
        else:
            self.good_thread = True

    def add_message(self, message):
        """
        Adds message to a thread and evaluates all thread attributes making changes as necessary.
        :param message: Message
        :return: None
        """
        self.message_count += 1
        if not self.good_thread and message.counts:
            # TODO This prevents single emails from support to a member/ideas etc. from being counted?
            self.good_thread = True

        new_stats, new_members = message.extract_labels()
        for label in new_stats:
            self.stat_labels.add(label)

        for label in new_members:
            self.member_labels.add(label)

        new_date = message.date
        if message.counts:
            if self.oldest_date is None or new_date < self.oldest_date:
                self.oldest_date = new_date
            if self.last_contact_date is None or new_date > self.last_contact_date:
                self.last_contact_date = new_date

        self.non_ping = len(self.stat_labels) > 0

        self._evaluate(message)

    def dont_count(self):
        self.good_thread = False

    def get_oldest_date(self):
        return self.oldest_date

    def get_last_contact_date(self):
        return self.last_contact_date

    def get_check_in_date(self):
        return self.check_in_date

    def get_count(self):
        return self.message_count

    def get_id(self):
        return self.id

    def get_stats(self):
        return list(self.stat_labels)

    def get_members(self):
        return list(self.member_labels)

    def get_subject(self):
        return self.subject

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

    def is_new_org(self):
        return self.new_org

    def is_closed(self):
        return self.closed

    def is_sales_ping(self):
        return self.sales_ping

    def is_check_in(self):
        return self.check_in_date is not None

    def __str__(self):
        return "Thread ID: " + self.id + ", Subject: " + self.subject() + ", Count: " + str(self.message_count)


class OpenInquiry:
    """
    An open thread in the Support Inbox.
    Attributes:
        id : int GMail thread id
        subject : str Thread subject line
    """
    def __init__(self, thread_id, subject):
        """
        Constructs a new open inquiry
        :param thread_id: : int GMail thread id
        :param subject: str Thread subject line
        """
        self.id = thread_id
        self.subject = subject

    def __repr__(self):
        return "< " + str(self.id) + ", " + self.subject + ">"

    def __eq__(self, other):
        return isinstance(other, OpenInquiry) and self.id == other.id

    def __hash__(self):
        return hash(self.id)

    @staticmethod
    def from_file(filename):
        """
        Builds a dictionary of open inquires from a file containing a list of open inquires.
        :param filename: File containing a list of open inquiries. Must be formatted properly
                         THREAD_ID
                         SUBJECT
                         No error is raised is file not formatted properly
        :return: dictionary of all open inquires {thread_id : OpenInquiry}
        :raise: IOError if the file cannot be read
        """
        file_in = open(filename, 'r')
        print "from file"
        threads = {}
        while file_in:
            print "reading"
            current = file_in.readline()
            if current is None:  # TODO Better way?
                break
            thread_id = file_in.readline().strip()
            subject = file_in.readline().strip()
            threads[thread_id] = OpenInquiry(thread_id, subject)
            break
        file_in.close()
        return threads

    @staticmethod
    def _from_message_list(messages):
        """
        Constructs a new open inquiry dictionary from a list of Messages
        :param messages: lst(Message)
            Constructs an OpenInquiry for each message in the list.
        :return: dictionary of all open inquiries {thread_id : OpenInquiry}
        """
        inbox = {}
        for message in messages:
            thread_id = message['X-GM-THRID']
            subject = message['Subject']

            if any(x is None for x in [thread_id, subject]):
                pass
            else:
                inbox[thread_id] = subject

        return inbox

    @staticmethod
    def from_current_inbox(inbox):
        """
        Builds a dictionary of open inquiries from a list of GMail messages.
        :param inbox: A list of GMail messages. Not a list of mail.Message types.
        :return: {Thread ID: OpenInquiry} for currently open and good threads.
        """
        threads = {}
        for msg in inbox:
            message = Message(msg)
            thread_id = message.get_thread_id()
            if thread_id in threads:
                if threads[thread_id].is_good():
                    threads[thread_id].add_message(message)
            else:
                threads[thread_id] = Thread(message)

            if not threads[thread_id].checked:
                if message.is_to_from_support():
                    threads[thread_id].should_it_count(message, 'to and from Support', True)
                elif message.is_internal():
                    threads[thread_id].should_it_count(message, 'internal', True)

            for l in config.OPEN_LABELS:
                if any(l in label for label in message.get_labels()):
                    threads[thread_id].closed = False

        open_inquiries = {}
        for thread in threads:
            trd = threads[thread]
            if trd.is_good() and not trd.is_closed():
                open_inquiries[trd.get_id()] = OpenInquiry(trd.get_id(), trd.get_subject())

        return open_inquiries

    @staticmethod
    def update(open_inquiries, new_open_inquiries, inbox):
        """
        Updates the current open_inquires with any new open inquires and removes those which have been closed.
        Calculates the number of inquires which remained open and the number which were closed.

        Writes the current list of open inquires to 'config.open.txt' which has the following format
            THREAD_ID
            SUBJECT

        :param open_inquiries: dict(OpenInquiry)
            Existing open inquiries
        :param new_open_inquiries: dict(OpenInquiry)
            Open inquires which which will be added to open_inquiries
        :param inbox: dict(OpenInquiry)
            Inquires currently in the inbox
        :return: (num_open : int, num_closed : int)
            num_open = Number of inquires in open_inquiries that are also in inbox
            num_closed = Number of inquires in open_inquiries that are not in inbox
        """
        num_open = 0
        num_closed = 0
        to_delete = []
        current = OpenInquiry._from_message_list(inbox)

        for thread in open_inquiries:
            if open_inquiries[thread].id in current:
                num_open += 1
            else:
                num_closed += 1
                to_delete.append(thread)

        for thread in to_delete:
            del open_inquiries[thread]

        OpenInquiry._write_to_file(open_inquiries.values(), 'config/open.txt')

        open_inquiries.update(new_open_inquiries)

        return num_open, num_closed

    @staticmethod
    def _write_to_file(open_inquiries, filename):
        """
         Writes the current list of open inquires to filename which has the following format
            THREAD_ID
            SUBJECT
        :param open_inquiries: lst(OpenInquiry)
        :param filename: destination file to which inquiries are written
        :raise IOError: If filename cannot be written to.
        :return: None
        """
        try:
            out = open(filename, 'wb')

            print "Recording open inquiries...\n"
            for thread in open_inquiries:
                out.write(thread.id + '\n')
                out.write(thread.subject + '\n')

            out.close()
        except IOError, e:
            print e
            print 'Could not open ' + filename + ' for writing.'
