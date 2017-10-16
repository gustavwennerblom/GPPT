

import exchangelib as exc
import exchangelib.ewsdatetime as ews
from exchangelib.attachments import FileAttachment, AttachmentId
import logging
from datetime import datetime
import random

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

        # timestamps the message with current date and time
        time = self.convert_to_EWStime(datetime.now())

        self.m = exc.Message(
            subject="[AM] Test",
            item_id = "DEF568",
            sender="Foo@Bar",
            body="Test message",
            datetime_sent = time
        )
        logging.info('Test message created')
        import os
        cwd=os.getcwd()
        filenames=os.listdir(cwd+u"/calcs")
        myfiles=[filenames[0]]
        for myfile in myfiles:
            try:
                os.chdir("./calcs")
                logging.info("Accessing: %s" % myfile)
                f = open(myfile, mode='rb')
                file_contents = f.read()
                suffix = random.randint(0,99999)
                att = FileAttachment(name=myfile, content=file_contents, attachment_id=AttachmentId())
                self.m.attach(att)
                f.close()
                os.chdir("..")
            except IOError:
                logging.error("Error opening file %s. Exiting." % myfile)
                import sys
                sys.exit(0)



