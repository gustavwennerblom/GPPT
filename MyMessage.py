

import exchangelib as exc
import exchangelib.ewsdatetime as ews
import datetime


class MyMessage:
    # Instancing the class returns a default message
    def get_message(self):
        return self.m

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

        time = ews.EWSDateTime(2017, 03, 02, 23, 29)

        self.m = exc.Message(
            subject="[AM] Test",
            sender="Foo@Bar",
            body="Test message",
            datetime_sent = time
        )

        import os
        cwd=os.getcwd()
        filenames=os.listdir(cwd+u"/calcs")
        myfile=filenames[1]
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



