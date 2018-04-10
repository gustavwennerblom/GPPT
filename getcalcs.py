import json
import config
import logging
import sys

from DBhelper_sqla import DBHelper, DuplicateFileWarning, DuplicateMessageWarning
from exchangelib import Account, Credentials, DELEGATE, Configuration, EWSDateTime
from excel_parser import ExcelParser, ExcelParsingError
from sqlalchemy.orm.exc import NoResultFound

# Initialize database manager script
db = DBHelper()   # Migrating away from one "db" class for all operations to avoid timeouts

# Initialize log
FORMAT = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
logging.basicConfig(filename="getcalcslog.log", format=FORMAT, level=config.loglevel)
log = logging.getLogger(__name__)

# Counts number of submissions in a given folder
def count_submissions_by_region(folder_name, account):
    f = account.inbox.get_folder_by_name(folder_name)
    return f.total_count


# Method to check a list of attachments and return the indexes of those eligible for analysis and storage
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
                logging.info(
                    'Attachment "%s" skipped, not in format for storage in database' % attachments[i].name
                )
        except UnicodeEncodeError as error:
            logging.error("Error caught in method 'check_attachments': " + repr(error))
    return indices


# Stores a submission received as an email message into the database. Returns the database index of that submission
def store_submission(mess):
    # db = DBHelper()
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

    # db.close()
    # Resetting return variable
    new_insert_indices = []
    for i in range(0, attachments_no):
        logging.info("Inserting submission message with subject: %s" % mess.subject)
        # creating new DBHelper to avoid timeouts
        # db = DBHelper()
        # db.insert_index returns the row id of the latest insert
        insert_index = db.insert_message(mess.attachments[attachment_indices[i]].name,
                                         mess.sender.email_address,
                                         mess.subject,
                                         str(mess.datetime_received),
                                         str(mess.item_id),
                                         str(mess.attachments[attachment_indices[i]].attachment_id),
                                         mess.attachments[attachment_indices[i]].content)

        # The database row of each insert (can be multiple if multiple eligible attachments is saved in a list)
        new_insert_indices.append(insert_index)
        # db.close()
    return new_insert_indices


# Returns list of emails received since last update of the database
def get_new_messages(folder_name, account, from_datetime='2018-01-01-00-00-00'):
    target_folder = account.inbox.get_folder_by_name(folder_name)
    tz = account.default_timezone
    from_datetime_list = [int(x) for x in from_datetime.split('-')]
    localized_from_datetime = tz.localize(EWSDateTime(*from_datetime_list))
    new_submissions = []
    for submission in target_folder.filter(datetime_received__gt=localized_from_datetime):
        new_submissions.append(submission)

    return new_submissions


# Looks through all messages in all folders in an account, finds new messages
def get_all_new_messages(account):
    # db = DBHelper()
    logging.info("Initializing message downloads")
    all_new_rows = []
    inbox = account.inbox
    # allfolders = account.inbox.children
    # allfolders.append(account.inbox)      # Breaks down in exchangelib 1.10
    allfolders = account.inbox.get_folders()

    last_update_timestamp = db.get_timestamp()
    tz = account.default_timezone
    from_datetime_list = [int(x) for x in last_update_timestamp.split('-')]
    localized_from_datetime = tz.localize(EWSDateTime(*from_datetime_list))

    for folder in allfolders:
        logging.info("Looking in folder %s" % str(folder))

        number_of_emails = folder.filter(datetime_received__gt=localized_from_datetime).count()
        logging.info("Found {0} new messages in folder {1}".format(number_of_emails, folder))
        for submission in folder.filter(datetime_received__gt=localized_from_datetime):
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

    db.set_timestamp()
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
    # db = DBHelper()
    tempfile = db.get_file_by_id(db_id)
    parser = ExcelParser(tempfile)

    try:
        db.insert_analysis(db_id, lead_office=parser.get_lead_office(),
                           project_margin=parser.get_margin(),
                           total_fee=parser.get_project_fee(),
                           total_hours=parser.get_total_hours(),
                           hours_by_role=parser.get_hours_by_role(),
                           blended_hourly_rate=parser.get_blended_hourly_rate(),
                           pricing_method=parser.assess_pricing_method(db_id),
                           tool_version=parser.determine_version())
    except TypeError:
        log.error("File on index {} not possible to parse. Skipping".format(db_id))
    #db.close()


# Fallback method to rerun analysis on all submissions stored in database, in case of failure half way
def reanalyze_all():
    # db = DBHelper()
    db_lines = db.get_all_ids()  # type: int
    log.info("Found {} lines in database".format(len(db_lines)))
    for index in db_lines:
        try:
            analyze_submission(index)
        except NoResultFound:
            logging.warning("analyze_submission failed on db index {0}. Possibly skip in index at write?".format(index))
    #db.close()


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
        logging.info("{} new rows inserted in database. Proceeding to analysis".format(len(all_new_rows)))

    # Loop iterates through all new inserted binaries and adds the excel analysis to the db
    for submission_id in all_new_rows:
        try:
            analyze_submission(submission_id)
        except ExcelParsingError as error:
            logging.error(repr(error))

    # Finally, reporting back all successful and closing database connection
    logging.info("All messages interpreted.")

def run_with_start_menu():
    user_select = input("Menu:\n "
                        "[1] Update all submissions \n "
                        "[2] Rerun quote analysis on database \n "
                        "[3] Set last updated timestamp to 2018-03-01"
                        "[5] Get list of messages new since last database update \n "
                        "[0] Devtest \n "
                        ">>")
    if user_select == "1":
        main()
    elif user_select == "2":
        logging.info("Initializing database re-analysis")
        reanalyze_all()
    elif user_select == "3":
        timestamp = '2018-03-01-00-00-00'
        db.set_timestamp(timestamp)
    elif user_select == '5':
        with open("./credentials/CREDENTIALS.json") as j:
            text = j.readline()
            d = json.loads(text)
        credentials = Credentials(username=d["UID"], password=d["PWD"])
        logging.info("Accessing Exchange credentials")
        config_office365 = Configuration(server="outlook.office365.com", credentials=credentials)
        account = Account(primary_smtp_address="projectproposals@business-sweden.se", config=config_office365,
                          autodiscover=False, access_type=DELEGATE)
        latest_database_update = db.get_timestamp()
        print('Database latest updated ' + latest_database_update)
        new_messages = get_new_messages("Americas", account, from_datetime=latest_database_update)
        for message in new_messages:
            print("Subject: {} Sent: {}".format(message.subject, message.datetime_sent))
        print("All done. Bye")
    else:
        print(user_select)
        print('Invalid selection. Please restart and try again')


def run_as_service(hours_between_runs):
    while True:
        main()
        time.sleep(60*60*hours_between_runs)


if __name__ == '__main__':
    if len(sys.argv) == 1:
        run_as_service(24)
    if len(sys.argv) == 2:
        if sys.argv[1] == '-menu':
            run_with_start_menu()
    else:
        print("Too many arguments")
        sys.exit(0)




