import json
import config
import logging

from DBhelper import DBHelper, DuplicateFileWarning, DuplicateMessageWarning
from exchangelib import Account, Credentials, DELEGATE, Configuration
from excel_parser import ExcelParser, ExcelParsingError

# Initialize database manager script
# db = DBHelper()   # Migrating away from one "db" class for all operations to avoid timeouts

# Initialize log
FORMAT = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
logging.basicConfig(filename="getcalcslog.log", format=FORMAT, level=config.loglevel)


# Counts number of submissions in a given folder
def count_submissions_by_region(folder_name, account):
    f = account.inbox.get_folder_by_name(folder_name)
    return f.total_count


# Method to check a list of attachments and return the indexes of those eiligible for analysis and storage
def check_attachments(attachments):
    # assert (attachments, exchangelib.folders.Attachment)
    rbound = len(attachments)
    indices = []
    for i in range(0, rbound):
        try:
            if attachments[i].name[-4:] == "xlsm":
                indices.append(i)
                logging.info('Attachment %s queued for storage in database' % attachments[i].name)
            else:
                logging.warning(
                    'Attachment "%s" skipped, not in format for storage in database' % attachments[i].name
                )
        except UnicodeEncodeError as error:
            logging.error("Error caught in method 'check_attachments': " + repr(error))
    return indices


# Stores a submission received as an email message into the database. Returns the database index of that submission
def store_submission(mess):
    db = DBHelper()
    # Checking if this specific Message has been store in the database already. Return 'None' if duplicate.
    if db.duplicatemessage(mess) and config.enforce_unique_messages:
        logging.warning('Attempt to insert message with EWS ID %s disallowed by configuration. '
                        'Set enforce_unique_files in config.py to "False" to allow.' % mess.item_id)
        try:
            raise DuplicateMessageWarning('Message (subject "%s") with this EWS message ID is already '
                                        'in database' % mess.subject)
        except UnicodeEncodeError as error:
            logging.error("Message with this EWS message ID (and cumbersome Unicode title) "
                          "already in database " + repr(error))

    # Check for eligible attachments and stores references to those in the attachment list
    attachment_indices = check_attachments(mess.attachments)
    attachments_no = len(attachment_indices)
    logging.info('Found %i attachments' % attachments_no)

    db.set_timestamp()
    db.close()
    # Resetting return variable
    new_insert_indices = []
    for i in range(0, attachments_no):

        logging.info("Inserting submission message with subject: %s" % mess.subject)
        # creating new DBHelper to avoid timeouts
        db = DBHelper()
        # db.insert_message returns the row id of the latest insert
        insert_index = db.insert_message(mess.attachments[attachment_indices[i]].name,
                                         mess.sender.email_address,
                                         mess.subject,
                                         str(mess.datetime_received),
                                         str(mess.item_id),
                                         str(mess.attachments[attachment_indices[i]].attachment_id),
                                         mess.attachments[attachment_indices[i]].content)

        # The database row of each insert (can be multiple if multiple eligible attachments is saved in a list
        new_insert_indices.append(insert_index)
        db.close()
    return new_insert_indices


# Returns list of emails received since last update of the database
def get_new_messages(folder_name, account):
    target_folder = account.inbox.get_folder_by_name(folder_name)

    new_submissions = []
    for submission in target_folder.all():

        new_submissions.append(submission)

    return new_submissions


# Looks through all messages in all folders in an account, finds new messages
def get_all_new_messages(account):
    logging.info("Initializing message downloads")
    all_new_rows = []
    allfolders = account.inbox.get_folders()
    # allfolders.append(account.inbox)      # Breaks down in exchangelib 1.10

    for folder in allfolders:
        logging.info("Looking in folder %s" % str(folder))

        # # Selection set for only looking in specific folders
        # if str(folder) in ["Messages (Americas)", "Messages (APAC)"]:
        #     logging.warning("Folder {0} skipped due to bugfix".format(folder))
        #     continue

        number_of_emails = folder.all().count()
        logging.info("Found {0} messages in folder {1}".format(number_of_emails, folder))
        for submission in folder.all().iterator():
            try:
                logging.info('Accessing submission with subject "%s"' % submission.subject)
                db_indices = store_submission(submission)
                # Unless store_submission has returned None, save row id in list to analyze
                if isinstance(db_indices, list):
                    all_new_rows += db_indices
            except DuplicateFileWarning as error:
                logging.warning(repr(error))
            except DuplicateMessageWarning as error:
                logging.warning(repr(error))
    return all_new_rows


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
    for region in regions:
        print("%s: %i submissions" % (region, count_submissions_by_region(region, account)))


# Reviews an xlsm file in the database, parses its contents, and commits the content to the database
def analyze_submission(db_id):
    db = DBHelper()
    tempfile = db.get_file_by_id(db_id)
    parser = ExcelParser(tempfile)

    db.insert_analysis(db_id, lead_office=parser.get_lead_office(),
                       project_margin=parser.get_margin(),
                       total_fee=parser.get_project_fee(),
                       total_hours=parser.get_total_hours(),
                       hours_by_role=parser.get_hours_by_role(),
                       blended_hourly_rate=parser.get_blended_hourly_rate(),
                       pricing_method=parser.assess_pricing_method(db_id),
                       tool_version=parser.determine_version())

    db.close()


# Fallback method to rerun analysis on all submissions stored in database, in case of failure half way
def reanalyze_all():
    db = DBHelper()
    db_lines = db.countlines()  # type: int
    for index in range(1, db_lines + 1):
        try:
            analyze_submission(index)
        except TypeError:
            logging.warning("analyze_submission failed on db index {0}. Possibly skip in index at write?".format(index))
    db.close()


def main():
    # Credentials for access to mailbox
    # Ignore if in debug mode when working with a spoof message
    credentials = None
    if not config.debug:
        with open("./credentials/CREDENTIALS.json") as j:
            text = j.readline()
            d = json.loads(text)
        credentials = Credentials(username=d["UID"], password=d["PWD"])
        logging.info("Accessing Exchange credentials")

    # Referencing Exchange account to fetch submissions from the projectproposal mailbox
    # Ignore if in debug mode when working with a spoof message
    if not config.debug:
        config_office365 = Configuration(server="outlook.office365.com", credentials=credentials)
        # noinspection PyUnboundLocalVariable
        account = Account(primary_smtp_address="projectproposals@business-sweden.se", config=config_office365,
                          autodiscover=False, access_type=DELEGATE)

    if config.debug:
        # USE BELOW TO INSERT A TEST MESSAGE IN THE DATABASE (CONFIGURE TEST MESSAGE IN MyMessage.py
        logging.info('Triggering import of test Message')
        from MyMessage import MyMessage
        m = MyMessage()
        testmail = m.get_message()
        all_new_rows = []

        try:
            all_new_rows = store_submission(testmail)
        except DuplicateFileWarning as error:
            logging.error(repr(error))
        except DuplicateMessageWarning as error:
            logging.error(repr(error))

    else:
        logging.info("Entering live mode with connection to Exchange server")
        # noinspection PyUnboundLocalVariable
        # messages = get_new_messages("Americas", account)  # Early test of mailbox access
        # Triggers run on all exchange folders and stores messages in db
        all_new_rows = get_all_new_messages(account)

    # Loop iterates through all new inserted binaries and adds the excel analysis to the db
    for submission_id in all_new_rows:
        try:
            analyze_submission(submission_id)
        except ExcelParsingError as error:
            logging.error(repr(error))

    # Finally, reporting back all sucessful and closing database connection
    logging.info("All messages interpreted.")

if __name__ == '__main__':
    user_select = input("Menu:\n "
                        "[1] Update all submissions \n "
                        "[2] Rerun quote analysis on database \n "
                        "[3] Export to csv file \n "
                        "[4] Get filename of first entry in GPPT_Submissions \n"
                        ">>")
    if user_select == "1":
        main()
    elif user_select == "2":
        logging.info("Initializing database re-analysis")
        reanalyze_all()
    elif user_select == "3":
        logging.info("Initializing export")
        db_main = DBHelper()
        db_main.export_db(format="csv")
        db_main.close()
    elif user_select == "4":
        db_main = DBHelper()
        stmt = "SELECT (Filename) FROM gppt_submissions WHERE (ID=%s)"
        db_main.cur.execute(stmt, (2,))
        r2 = db_main.cur.fetchone()[0]
        db_main.close()
        print(r2)
        print("Bye")
    else:
        print(user_select)
        print('Valid selections only "1" or "2" or "3". Please restart and try again')



