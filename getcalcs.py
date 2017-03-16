

import json
import config
import logging

from DBhelper import DBHelper, DuplicateFileError, DuplicateMessageError
from exchangelib import Account, Credentials, DELEGATE
from excel_parser import ExcelParser
import exchangelib.ewsdatetime
from datetime import datetime

db = DBHelper()
logging.basicConfig(filename="getcalcslog.log", level = config.loglevel)

# Counts number of submissions in a given folder
def count_submissions_by_region(folder_name, account):
    f = account.inbox.get_folder_by_name(folder_name)
    # counter = 0
    # for item in f.all():
    #     counter += 1
    # return counter
    return f.total_count

# Method to check a list of attachments and return the indexes of those eiligible for analysis and storage
def check_attachments(attachments):
    # assert (attachments, exchangelib.folders.Attachment)
    rbound = len(attachments)
    indices = []
    for i in range(0, rbound):
        if attachments[i].name[-4:] == "xlsm":
            indices.append(i)
        else:
            logging.warning(
                'Attachment "{}" skipped, not in format for storage in database'.format(attachments[i].name))
    return indices

# Stores a submission received as an email message into the database. Returns the database index of that submission
def store_submission(mess):
    # Checking if this specific Message has been store in the database already
    if db.duplicatemessage(mess) and config.enforce_unique_messages:
        logging.warning('Attempt to insert message with EWS ID {} disallowed by configuration. Set enforce_unique_files '
                        'in config.py to "False" to allow.'.format(mess.item_id))
        raise TypeError('Message (subject "{}") with this EWS message ID is already in database'.format(mess.subject))


    # Check for eligible attachments and stores references to those in the attachment list
    attachment_indices = check_attachments(mess.attachments)
    logging.info('Found attachments {}'.format(mess.attachments))
    attachments_no = len(attachment_indices)
    logging.info('Proceeding to store attachment number {0} in sequence of {1}.'.format(attachment_indices, attachments_no))

    db.set_timestamp()


    for i in range(0, attachments_no):
        logging.info("Inserting submission message with subject: %s" % mess.subject)
        insert_index = db.insert_message(mess.attachments[attachment_indices[i]].name,
                                         mess.sender.email_address,
                                         mess.subject,
                                         str(mess.datetime_received),
                                         str(mess.item_id),
                                         str(mess.attachments[attachment_indices[i]].attachment_id),
                                         mess.attachments[attachment_indices[i]].content)
        if insert_index:
            analyze_submission(insert_index)

    return mess.item_id

# Returns list of emails received since last update of the database
def get_new_messages(folder_name, account):
    target_folder = account.inbox.get_folder_by_name(folder_name)

    new_submissions=[]
    for submission in target_folder.all():

        new_submissions.append(submission)

    return new_submissions

# Test method to try attachment download
def get_one_message(folder_name, account):
    f = account.inbox.get_folder_by_name(folder_name)
    all_items = []
    for item in f.all().order_by('-datetime_received')[:10]:
        all_items.append(item)

    return all_items[0]

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
        # noinspection PyUnboundLocalVariable
        account = Account(primary_smtp_address="projectproposals@business-sweden.se", credentials=credentials,
                          autodiscover=True, access_type=DELEGATE)

    if config.debug:
        # USE BELOW TO INSERT A TEST MESSAGE IN THE DATABASE (CONFIGURE TEST MESSAGE IN MyMessage.py
        # insert_testmessage()
        logging.info('Triggering import of test Message')
        from MyMessage import MyMessage
        m = MyMessage()
        mess = m.get_message()
        messages = [mess]
    else:
        logging.info("Entering live mode with connection to Exchange server")
        messages = get_new_messages("Americas", account)    # This should be changed to go through all regional mailboxes


    # Three lines to first insert the file (with metadata), analyze the file, and close the connection
    for mess in messages:
        try:
            store_submission(mess)
        # except ValueError as error:
        #    logging.error(repr(error))
        # except TypeError as error:
        #    logging.error(repr(error))
        except DuplicateFileError as error:
            logging.error(repr(error))
        except DuplicateMessageError as error:
            logging.error(repr(error))

    logging.info("All messages stored and interpreted.")
    print ("All done, closing connection to database")
    logging.info("Closing database")
    db.close()

if __name__ == '__main__':
    main()
