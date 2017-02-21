import json
import sqlite3
import sys

from exchangelib import Account, Credentials, DELEGATE
from exchangelib.folders import FileAttachment

# Credentials for access to mailbox
with open("CREDENTIALS.json") as f:
    text=f.readline()
    d=json.loads(text)

credentials = Credentials(username=d["UID"], password=d["PWD"])

# Referencing Exchange account to fetch submissions from the projectproposal mailbox
account = Account(primary_smtp_address="projectproposals@business-sweden.se", credentials=credentials, autodiscover=True, access_type=DELEGATE)

# Obtain connection to database
try:
    DB = sqlite3.connect("submissions.db")
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
    f=get_folder(folder_name)
    counter=0
    for item in f.all():
        counter+=1
    return counter

def store_attachment(filename, submitter, region, date, checksum, file):
    #with DB:
    #    cur=DB.cursor()
    #    cur.execute("INSERT INTO ")
    pass

# Downloads all attachments into a database, checking for duplicates
def download_attachments(folder_name):
    pass

# Downloads an attachement into database given a subject line
def download_attachment(folder_name, subject):
    pass

# Generates a checksum for the file
def generate_hash(file):
    pass

#Basic run function to print a count of submissions by region
regions = ["Americas", "APAC", "MEA", "CEE", "Western Europe"]
for x in regions:
    print "%s: %i submissions" % (x, count_submissions_by_region(x))









