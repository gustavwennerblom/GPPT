import exchangelib as exc
import exchangelib.ewsdatetime as ews
from exchangelib.attachments import FileAttachment, AttachmentId
import logging
from datetime import datetime
import random
from collections import namedtuple

class MyMessage:
    # Instancing the class returns a default message
    def get_message(self):
        return self.m

    # Converts a datetime object to EWStime (as EWS wants it)
    def convert_to_EWStime(self, in_time):
        timestring = datetime.strftime(in_time, "%Y-%m-%d-%H-%M")
        t = timestring.split("-")
        return ews.EWSDateTime(int(t[0]), int(t[1]), int(t[2]), int(t[3]), int(t[4]))

    def __init__(self):
        # Get a logger
        log = logging.getLogger(__name__)

        # timestamps the message with current date and time
        time = self.convert_to_EWStime(datetime.now())

        # Create a namedtuple to simulate exchangelib Mailbox object
        FakeSender = namedtuple("FakeSender", ["email_address"])
        my_sender = FakeSender("foo@bar.com")

        self.m = exc.Message(
            subject="[AM] Test",
            item_id = "DEF568",
            sender=my_sender,
            body="Test message",
            datetime_sent = time
        )
        log.info('Test message created')
        import os
        cwd = os.path.dirname(os.path.realpath(__file__))
        # filenames = os.listdir(cwd+u"/calcs")
        dummy = os.path.join(cwd, 'calcs', 'Dummy.txt')
        actual = os.path.join(cwd, 'calcs', 'test.xlsx')
        myfiles = [actual]
        for myfile in myfiles:
            try:
                # os.chdir("./calcs")
                log.info("Accessing: %s" % myfile)
                f = open(myfile, mode='rb')
                file_contents = f.read()
                # suffix = random.randint(0,99999)
                att = FileAttachment(name=myfile, content=file_contents, attachment_id=AttachmentId())
                self.m.attach(att)
                f.close()
                os.chdir("..")
            except IOError:
                log.error("Error opening file %s. Exiting." % myfile)
                import sys
                sys.exit(0)



