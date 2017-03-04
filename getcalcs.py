

import json
import config

from DBhelper import DBHelper
from exchangelib import Account, Credentials, DELEGATE
from excel_parser import ExcelParser

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

db = DBHelper()

# Function to return a folder given a folder name
def get_folder(name):
    return account.inbox.get_folder_by_name(name)


# Counts number of submissions in a given folder
def count_submissions_by_region(folder_name):
    f = get_folder(folder_name)
    counter = 0
    for item in f.all():
        counter += 1
    return counter


def store_attachment(filename, submitter, region, date, checksum, attachment):
    # with DB:
    #    cur=DB.cursor()
    #    cur.execute("INSERT INTO ")
    pass


# Downloads all attachments in a folder into a database, checking for duplicates
def download_attachments(folder_name):
    # f = get_folder(folder_name)
    # for item in f.all():
    #    store_attachment(item.)
    pass


# Test method to try attachment download
def download_one_attachment(folder_name):
    f = get_folder(folder_name)
    all_items = []
    for item in f.all():
        all_items.append(item)

    #Drop to shell to test code
    import code
    code.interact(local=locals())

    # filename = all_items[0].attachments[0].name
    # submitter = all_items[0].author.email_address
    # region = all_items[0].subject[:4]
    # date = all_items[0].datetime_received  # need to find the __tostring__ function
    # attachment_id = all_items[0].attachments[0].attachment_id



# Downloads an attachment into database given a subject line
def download_attachment(folder_name, subject):
    pass


# Generates a checksum for the file
def generate_hash(attachment):
    pass


# Basic run function to print a count of submissions by region
def count_all():
    regions = ["Americas", "APAC", "MEA", "CEE", "Western Europe"]
    for x in regions:
        print "%s: %i submissions" % (x, count_submissions_by_region(x))

# Subrouting for testing. Inserts a spoof message into the database
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

        db.insert(testmail.attachments[0].name, testmail.sender, testmail.subject, str(testmail.datetime_sent), str(testmail.attachments[0].attachment_id), testmail.attachments[0].content)
        db.close()
    else:
        print "Subroutine not intended for production use. Set debug=True in config.py to use"

# Reviews an xlsm file in the database and prints to stdout a set of data about it
def analyze_submission(db_id):
    tempfile=db.get_file_by_id(db_id)
    print 'Temporary file "%s" created' % tempfile
    parser=ExcelParser(tempfile)
    print "Lead office: %s" % parser.get_lead_office()
    print "Project margin: %s" % parser.get_margin()
    print "Total fee: %s" % parser.get_project_fee()
    print "Total hours: %s" % parser.get_total_hours()
    print "By role: "
    print parser.get_hours_by_role()
    print "Blended hourly rate: %s" % parser.get_blended_hourly_rate()
    print "Pricing method: %s" % parser.assess_pricing_method()
    # Drop to shell to test code
    #import code
    #code.interact(local=locals())



if config.debug:
    #USE BELOW TO INSERT A TEST MESSAGE IN THE DATABASE (CONFIGURE TEST MESSAGE IN MyMessage.py
    #insert_testmessage()

    #USE BELOW TO CHECK XLSM ANALYTICS
    #analyze_submission(18)

else:
    print("Entering live mode with connection to Exchange server")
    download_one_attachment("Americas")
    #print("No production lines present")
