import exchangelib as exc


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

        self.m = exc.Message(
            subject="Test",
            sender="Foo@Bar",
            body="<html><head></head><body><p>Test message</p></body></html>",
        )
        with open("Testfile.xlsm", mode='rb') as f:
            file_contents = f.read()
            att = exc.folders.FileAttachment(name="Testfile.xlsm", content=file_contents, attachment_id=exc.folders.AttachmentId("ABC1234"))
            self.m.attach(att)
