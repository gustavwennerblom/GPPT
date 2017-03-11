

import json
import config

from DBhelper import DBHelper
from exchangelib import Account, Credentials, DELEGATE
from excel_parser import ExcelParser
import exchangelib.ewsdatetime
from datetime import datetime
db = DBHelper()


# Counts number of submissions in a given folder
def count_submissions_by_region(folder_name, account):
    f = account.inbox.get_folder_by_name(folder_name)
    # counter = 0
    # for item in f.all():
    #     counter += 1
    # return counter
    return f.total_count

# Converts a DateTime object to EWSDataTime
# noinspection PyPep8Naming
# def convert_to_EWStime(in_time):
#     timestring = datetime.strftime(in_time, "%Y-%m-%d-%H-%M")
#     t = timestring.split("-")
#     return ews.EWSDateTime(int(t[0]), int(t[1]), int(t[2]), int(t[3]), int(t[4]))

# Stores a submission received as an email message into the database. Returns the database index of that submission
def store_submission(mess):
    if db.duplicatemessage(mess):
        raise TypeError('Message (subject "{}") with this EWS message ID is already in database'.format(mess.subject))

    print "Inserting submission message with subject: %s" % mess.subject


    db.update_timestamp()
    insert_index = db.insert_message(mess.attachments[0].name,
                                     mess.sender,
                                     mess.subject,
                                     str(mess.datetime_sent),
                                     str(mess.item_id),
                                     str(mess.attachments[0].attachment_id),
                                     mess.attachments[0].content)
    return insert_index

# Returns list of emails received since last update of the database
def get_new_messages(folder_name, account):
    target_folder = account.inbox.get_folder_by_name(folder_name)
    latest_update = db.get_timestamp()
    latest_update_ews = exchangelib.EWSDateTime.from_datetime(latest_update)
    ordered_submissions = target_folder.all().order_by('-datetime_received')
    new_submissions=[]
    for submission in ordered_submissions:
        if submission.datetime_sent > latest_update_ews:
            break
        new_submissions.append(submission)

    return new_submissions

# Test method to try attachment download
def get_one_message(folder_name, account):
    f = account.inbox.get_folder_by_name(folder_name)
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
        db.set_timestamp(datetime(2017,01,01))
        print("Resetting last update to 2017-01-01 to get all old messages (test)")
        messages = get_new_messages("Americas", account)

        import code
        code.interact(local=locals())
        # mess = get_one_message("Americas", account)
        # print("No production lines present")

    # Three lines to first insert the file (with metadata), analyze the file, and close the connection
    try:
        insert_index = store_submission(mess).fetchone()[0]
        analyze_submission(insert_index)
        db.close()
    except ValueError as error:
        print repr(error)
    except TypeError as error:
        print repr(error)

    # USE BELOW TO CHECK XLSM ANALYTICS
    # analyze_submission(18)

if __name__ == '__main__':
    main()
