import config
import logging
import sys
import os
import json
from logging.handlers import TimedRotatingFileHandler
from logging import StreamHandler
from datetime import datetime

from DBhelper_sqla import DBHelper, DuplicateFileWarning, DuplicateMessageWarning
from exchangelib import Account, Credentials, DELEGATE, Configuration, EWSDateTime
from excel_parser import ExcelParser, ExcelParsingError
from sqlalchemy.orm.exc import NoResultFound
from mailer import LogMailer

# Initialize database manager script. May need to be re-initialized on MySQL with aggressive timeout settings
db = DBHelper()

# Initialize and configure log
script_directory = os.path.dirname(os.path.realpath(__file__))
log = logging.getLogger("getcalcs-main")
log.setLevel(logging.INFO)

# Removed for webapp deployment
# ch1 = TimedRotatingFileHandler(filename=os.path.join(script_directory, config.log_directory, config.main_log_filename),
#                               when='d', interval=1, backupCount=7)
# ch1.setLevel(config.loglevel)
# ch1.setFormatter(logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s'))
# log.addHandler(ch1)

ch2 = StreamHandler(stream=sys.stderr)
ch2.setLevel(config.loglevel)
ch2.setFormatter(logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s'))
log.addHandler(ch2)
log.info("Main log set up")
log.debug("Debug message")
log.info("Another log message")

# Set severity level of exchangelibs logger
logging.getLogger('exchangelib').setLevel(logging.WARNING)


# Counts number of submissions in a given folder
def count_submissions_by_region(folder_name, account):
    f = account.inbox.get_folder_by_name(folder_name)
    return f.total_count


# Checks a list of attachments and returns the indices of those eligible for analysis and storage
def check_attachments(attachments):
    r_bound = len(attachments)
    indices = []
    for i in range(0, r_bound):
        try:
            if attachments[i].name[-4:] == "xlsm":
                indices.append(i)
                log.info('Attachment %s queued for storage in database' % attachments[i].name)
            else:
                log.info(
                    'Attachment "%s" skipped, not in format for storage in database' % attachments[i].name
                )
        except UnicodeEncodeError as error:
            log.error("Error caught in method 'check_attachments': " + repr(error))
    return indices


# Stores a submission received as an email message into the database. Returns the database index of that submission
def store_submission(mess):
    # Check if this specific Message has been store in the database already. Return 'None' if duplicate.
    if db.duplicatemessage(mess) and config.enforce_unique_messages:
        log.warning('Attempt to insert message with EWS ID %s disallowed by configuration. '
                    'Set enforce_unique_files in config.py to "False" to allow.' % mess.item_id)
        try:
            raise DuplicateMessageWarning('Message (subject "%s") with this EWS message ID is already '
                                          'in database' % mess.subject)
        except UnicodeEncodeError as error:
            log.error("Message with this EWS message ID (and cumbersome Unicode title) "
                      "already in database " + repr(error))

    # Check for eligible attachments and stores references to those in the attachment list
    attachment_indices = check_attachments(mess.attachments)
    attachments_no = len(attachment_indices)
    log.info('Found %i attachments' % attachments_no)

    # Resetting return variable
    new_insert_indices = []
    for i in range(0, attachments_no):
        log.info("Inserting submission message with subject: %s" % mess.subject)

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
def fetch_all_new_messages(account):
    log.info("Initializing message downloads")
    all_new_rows = []

    allfolders = account.inbox.get_folders()

    last_update_timestamp = db.get_timestamp()
    tz = account.default_timezone
    from_datetime_list = [int(x) for x in last_update_timestamp.split('-')]
    localized_from_datetime = tz.localize(EWSDateTime(*from_datetime_list))

    for folder in allfolders:
        log.info("Looking in folder %s" % str(folder))

        number_of_emails = folder.filter(datetime_received__gt=localized_from_datetime).count()
        log.info("Found {0} new messages in folder {1}".format(number_of_emails, folder))
        for submission in folder.filter(datetime_received__gt=localized_from_datetime):
            try:
                log.info('Accessing submission with subject "%s"' % submission.subject)
                db_indices = store_submission(submission)
                # Unless store_submission has returned None, save row id in list to analyze
                if isinstance(db_indices, list):
                    all_new_rows += db_indices
            except DuplicateFileWarning as error:
                log.warning(repr(error))
            except DuplicateMessageWarning as error:
                log.warning(repr(error))

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


# Fallback method to rerun analysis on all submissions stored in database, in case of failure half way
def reanalyze_all():
    db_lines = db.get_all_ids()  # type: list
    log.info("Found {} lines in database".format(len(db_lines)))
    for index in db_lines:
        try:
            analyze_submission(index)
        except NoResultFound:
            log.warning("analyze_submission failed on db index {0}. Possibly skip in index at write?".format(index))


def main():
    # Credentials for access to mailbox
    with open(os.path.join(script_directory, 'credentials', 'CREDENTIALS.json')) as j:
        text = j.readline()
        d = json.loads(text)
    credentials = Credentials(username=d["UID"], password=d["PWD"])
    log.info("Accessing Exchange credentials")

    # Referencing Exchange account to fetch submissions from the projectproposal mailbox
    config_office365 = Configuration(server="outlook.office365.com", credentials=credentials)
    account = Account(primary_smtp_address="projectproposals@business-sweden.se", config=config_office365,
                      autodiscover=False, access_type=DELEGATE)
    log.info("Entering live mode with connection to Exchange server")

    # Triggers run on all exchange folders and stores messages in db
    all_new_rows = fetch_all_new_messages(account)
    log.info("{} new rows inserted in database. Proceeding to analysis".format(len(all_new_rows)))

    # Loop iterates through all new inserted binaries and adds the excel analysis to the db
    if all_new_rows:
        for submission_id in all_new_rows:
            try:
                analyze_submission(submission_id)
            except ExcelParsingError as error:
                log.error(repr(error))
        log.info("{} new submissions analyzed (with or without errors).".format(len(all_new_rows)))
    else:
        log.info("No new inserts in database. No new analyses to do.")

    # Finally, sending log file as email - DISABLED FOR WEB APP
    # log.info("Commencing log email distribution.")
    # log_mailer = LogMailer()
    # timestamp = datetime.strftime(datetime.now(), "%Y-%m-%d %H:%M")
    # log_mailer.send_mail("GPPT get calcs executed",
    #                      "This message was autogenerated on {}".format(timestamp),
    #                      attachment_path=os.path.join(script_directory, config.log_directory, config.main_log_filename))


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
        log.info("Initializing database re-analysis")
        reanalyze_all()
    elif user_select == "3":
        timestamp = '2018-03-01-00-00-00'
        db.set_timestamp(timestamp)
    elif user_select == '5':
        with open(os.path.join(script_directory, 'credentials', 'CREDENTIALS.json')) as j:
            text = j.readline()
            d = json.loads(text)
        credentials = Credentials(username=d["UID"], password=d["PWD"])
        log.info("Accessing Exchange credentials")
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


if __name__ == '__main__':
    if len(sys.argv) == 1:
        print("Too few arguments. Use -menu or -update to run")
    if len(sys.argv) == 2:
        if sys.argv[1] == '-menu':
            run_with_start_menu()
        if sys.argv[1] == '-update':
            main()
    else:
        print("Too many arguments")
        sys.exit(0)




