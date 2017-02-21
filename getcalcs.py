import json

from exchangelib import Account, Credentials, DELEGATE
from exchangelib.folders import FileAttachment

# Credentials for access to mailbox
# TO BE MOVED TO SEPARATE JSON AND NOT CHECKED INTO GIT
with open("CREDENTIALS.json") as f:
    text=f.readline()
    d=json.loads(text)

credentials = Credentials(username=d["UID"], password=d["PWD"])

# Referencing Exchange account to fetch submissions from
account = Account(primary_smtp_address="projectproposals@business-sweden.se", credentials=credentials, autodiscover=True, access_type=DELEGATE)

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

#Basic run function to print a count of submissions by region
regions = ["Americas", "APAC", "MEA", "CEE", "Western Europe"]
for x in regions:
    print "%s: %i submissions" % (x, count_submissions_by_region(x))









