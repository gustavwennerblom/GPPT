

import exchangelib as exc
import exchangelib.ewsdatetime as ews
from datetime import datetime


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
        # credentials = exc.Credentials("foo", "bar", is_service_account=False)
        # config = exc.Configuration(
        #     server="bar.com",
        #     credentials=credentials
        # )
        #
        # self.a = exc.Account(
        #     primary_smtp_address="foo@bar.com",
        #     config=config,
        #     autodiscover=False,
        #     access_type=exc.DELEGATE
        # )

        # timestamps the message with current date and time
        time = self.convert_to_EWStime(datetime.now())

        self.m = exc.Message(
            subject="[AM] Test",
            item_id = "DEF567",
            sender="Foo@Bar",
            body="Test message",
            datetime_sent = time
        )

        import os
        cwd=os.getcwd()
        filenames=os.listdir(cwd+u"/calcs")
        myfile=filenames[2]
        try:
            os.chdir("./calcs")
            print("Accessing: %s" % myfile + str(type(myfile)))
            f=open(myfile, mode='rb')
            file_contents = f.read()
            att = exc.folders.FileAttachment(name=myfile, content=file_contents, attachment_id=exc.folders.AttachmentId("ABC1234"))
            self.m.attach(att)
            f.close()
            os.chdir("..")
        except IOError:
            print("Error opening file %s. Exiting." % myfile)
            import sys
            sys.exit(0)



