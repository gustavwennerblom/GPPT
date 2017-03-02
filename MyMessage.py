import exchangelib as exc
import exchangelib.ewsdatetime as ews
import datetime

class Account(object):
    pass


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
            body="<html><head></head><body><p>Test message</p></body></html>",
            datetime_sent = time
        )
        with open("Testfile.xlsm", mode='rb') as f:
            file_contents = f.read()
            att = exc.folders.FileAttachment(name="Testfile.xlsm", content=file_contents, attachment_id=exc.folders.AttachmentId("ABC1234"))
            self.m.attach(att)
