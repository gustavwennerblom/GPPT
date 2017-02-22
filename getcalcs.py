import json
import sqlite3
import sys

from exchangelib import Account, Credentials, DELEGATE
# from exchangelib.folders import FileAttachment

# Credentials for access to mailbox
with open("CREDENTIALS.json") as j:
    text = j.readline()
    d = json.loads(text)

credentials = Credentials(username=d["UID"], password=d["PWD"])

# Referencing Exchange account to fetch submissions from the projectproposal mailbox
account = Account(primary_smtp_address="projectproposals@business-sweden.se", credentials=credentials,
                  autodiscover=True, access_type=DELEGATE)

# Obtain connection to database
DB = sqlite3.connect("submissions.db")
try:
    cur = DB.cursor()
except DB.Error, e:
    print "Error accessing database: %s" % e.args[0]
    sys.exit(1)
finally:
    if DB:
        DB.close()


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

#Test method to try attachment download
def download_one_attachment(folder_name):
    f = get_folder(folder_name)
    all_items = []
    for item in f.all():
        all_items.append(item)
    # Drop to shell to test code
    #import code
    #code.interact(local=locals())
    filename = all_items[0].attachments[0].name
    submitter=all_items[0].author.email_address
    region=all_items[0].subject[:4]
    date=all_items[0].datetime_received #need to find the __tostring__ function
    attachment_id=all_items[0].attachments[0].attachment_id

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


# Execution lines
download_one_attachment("Americas")
