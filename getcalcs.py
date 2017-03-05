

import json
import config

from DBhelper import DBHelper
from exchangelib import Account, Credentials, DELEGATE
from excel_parser import ExcelParser

db = DBHelper()


# Function to return a folder given a folder name
def get_folder(name, account):
    return account.inbox.get_folder_by_name(name)


# Counts number of submissions in a given folder
def count_submissions_by_region(folder_name, account):
    f = get_folder(folder_name, account)
    counter = 0
    for item in f.all():
        counter += 1
    return counter


# Stores a submission received as an email message into the database. Returns the database index of that submission
def store_submission(mess):

    print "Inserting submission message with:"
    print "\t Subject: %s" % mess.subject
    print "\t File: %s" % mess.attachments[0].name
    print "\ From: %s" % mess.sender
    print "\t Date: %s" % mess.datetime_sent
    print "\t Id: %s" % mess.attachments[0].attachment_id

    insert_index = db.insert_message(mess.attachments[0].name,
                                     mess.sender,
                                     mess.subject,
                                     str(mess.datetime_sent),
                                     str(mess.attachments[0].attachment_id),
                                     mess.attachments[0].content)

    return insert_index


def store_analysis():
    pass


# Downloads all attachments in a folder into a database, checking for duplicates
# Possibly not needed? Overridden by 'store_attachment'?
def download_attachments(folder_name):
    # f = get_folder(folder_name)
    # for item in f.all():
    #    store_attachment(item.)
    pass


# Test method to try attachment download
def download_one_attachment(folder_name, account):
    f = get_folder(folder_name, account)
    all_items = []
    for item in f.all():
        all_items.append(item)

    return all_items[0]
    # filename = all_items[0].attachments[0].name
    # submitter = all_items[0].author.email_address
    # region = all_items[0].subject[:4]
    # date = all_items[0].datetime_received  # need to find the __tostring__ function
    # attachment_id = all_items[0].attachments[0].attachment_id


# Basic run function to print a count of submissions by region
def count_all(account):
    regions = ["Americas", "APAC", "MEA", "CEE", "Western Europe"]
    for x in regions:
        print "%s: %i submissions" % (x, count_submissions_by_region(x, account))


# Subrouting for testing. Inserts a spoof message into the database
# Overridden by store_attacment(mess)
def insert_testmessage():
    if config.debug:
        print("Entering test/debug mode")
        from MyMessage import MyMessage
        m = MyMessage()
        testmail = m.get_message()
        print "Subject: %s" % testmail.subject
        print "File: %s" % testmail.attachments[0].name
        print "From: %s" % testmail.sender
        print "Date: %s" % testmail.datetime_sent
        print "Id: %s" % testmail.attachments[0].attachment_id

        db.insert_message(testmail.attachments[0].name, testmail.sender, testmail.subject, str(testmail.datetime_sent),
                          str(testmail.attachments[0].attachment_id), testmail.attachments[0].content)
        db.close()
    else:
        print "Subroutine not intended for production use. Set debug=True in config.py to use"


# Reviews an xlsm file in the database and prints to stdout a set of data about it
# noinspection PyDictCreation
def analyze_submission(db_id):
    tempfile = db.get_file_by_id(db_id)
    parser = ExcelParser(tempfile)
    # print 'Temporary file "%s" created' % tempfile
    # print "Lead office: %s" % parser.get_lead_office()
    # print "Project margin: %s" % parser.get_margin()
    # print "Total fee: %s" % parser.get_project_fee()
    # print "Total hours: %s" % parser.get_total_hours()
    # print "By role: "
    # print parser.get_hours_by_role()
    # print "Blended hourly rate: %s" % parser.get_blended_hourly_rate()
    # print "Pricing method: %s" % parser.assess_pricing_method()

    # build dict to pass to database
    d = {}
    d["lead_office"] = parser.get_lead_office()
    d["project_margin"] = parser.get_margin()
    d["total_fee"] = parser.get_project_fee()
    d["total_hours"] = parser.get_total_hours()
    d["hours_by_role"] = parser.get_hours_by_role()  # Note: This returns a dict
    d["blended_hourly_rate"] = parser.get_blended_hourly_rate()
    d["pricing_method"] = parser.assess_pricing_method()

    db.insert_analysis(db_id, lead_office=parser.get_lead_office(),
                       project_margin=parser.get_margin(),
                       total_fee=parser.get_project_fee(),
                       total_hours=parser.get_total_hours(),
                       hours_by_role=parser.get_hours_by_role(),
                       blended_hourly_rate=parser.get_blended_hourly_rate(),
                       pricing_method=parser.assess_pricing_method())

    # Drop to shell to test code
    # import code
    # code.interact(local=locals())


def main():

    # Credentials for access to mailbox
    # Ignore if in debug mode when working with a spoof message
    if not config.debug:
        with open("CREDENTIALS.json") as j:
            text = j.readline()
            d = json.loads(text)
        credentials = Credentials(username=d["UID"], password=d["PWD"])

    # Referencing Exchange account to fetch submissions from the projectproposal mailbox
    # Ignore if in debug mode when working with a spoof message
    if not config.debug:
        account = Account(primary_smtp_address="projectproposals@business-sweden.se", credentials=credentials,
                          autodiscover=True, access_type=DELEGATE)

    if config.debug:
        # USE BELOW TO INSERT A TEST MESSAGE IN THE DATABASE (CONFIGURE TEST MESSAGE IN MyMessage.py
        # insert_testmessage()
        from MyMessage import MyMessage
        m = MyMessage()
        mess = m.get_message()
    else:
        print("Entering live mode with connection to Exchange server")
        mess = download_one_attachment("Americas", account)
        # print("No production lines present")

    # Three lines to first insert the file (with metadata), analyze the file, and close the connection
    insert_index = store_submission(mess).fetchone()[0]
    analyze_submission(insert_index)
    db.close()

    # USE BELOW TO CHECK XLSM ANALYTICS
    # analyze_submission(18)

if __name__ == '__main__':
    main()
