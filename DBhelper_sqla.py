import logging
import config
from datetime import datetime
import credentials.DBcreds as DBcreds
from sqlalchemy import create_engine, Table, String, Integer, MetaData
from sqlalchemy.orm import sessionmaker
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.sql.expression import func
from models import Base, LastUpdated

if config.test_database:
    from models import DevSubmission as Submission, DevLastUpdated as LastUpdated
else:
    from models import Submission, LastUpdated

log = logging.getLogger(__name__)
log.info("Logging initiated")


class DBHelper:

    # Sets a timestamp (string) in the 'latest updated' table in the database and returns it (as string)
    def set_timestamp(self, *args):
        if len(args) > 2:
            log.error('set_timestamp takes maximum one argument, %i given' % len(args))
            raise TypeError("set_timestamp in DBhelper misused. Check log.")
        elif len(args) == 1:
            timestamp = args[0]
        else:
            t = datetime.now()
            timestamp = datetime.strftime(t, "%Y-%m-%d-%H-%M-%S")

        res = self.session.query(LastUpdated).filter_by(Id=1)
        if not res.all():
            first_timestamp = LastUpdated(Updated=timestamp)
            self.session.add(first_timestamp)
        else:
            res.one().Updated = timestamp
        self.session.commit()
        log.info('Database timestamp for latest update set to %s' % timestamp)
        return timestamp

    # Fetches and returns the latest updated timestamp (string) from the database
    def get_timestamp(self):
        res = self.session.query(LastUpdated).one()
        log.info('Returning timestamp {}'.format(res.Updated))
        return res.Updated

    # Checks if a given filename is in the database. Returns boolean.
    def duplicate_file(self, filename):
        log.debug("Checking if file with name {} is already in database".format(filename))
        hits = self.session.query(Submission).filter_by(Filename=filename).all()
        if hits:
            return True
        else:
            return False

    # Checks if a given message is in the database. Returns boolean.
    def duplicatemessage(self, message):
        log.debug("Checking if message with EWS id {} is in already in database".format(message.item_id))
        hits = self.session.query(Submission).filter_by(Message_Id=message.item_id).all()
        return bool(hits)

    # Inserts an set of data as a new row in the database. Returns the index of that last insert. Returns None if dupe.
    def insert_message(self, filename, submitter, region, date, message_id, attachment_id, attachment):
        # Check if file is already in database else return None
        if self.duplicate_file(filename) and config.enforce_unique_files:
            log.info('Attempt to insert file {} in database disallowed. Set enforce_unique_files in '
                     'config.py to "False" to allow'.format(filename))
            return None

        new_submission = Submission(Filename=filename,
                                    Submitter=submitter,
                                    Region=region,
                                    Date=date,
                                    Message_Id=message_id,
                                    Attachment_Id=attachment_id,
                                    Attachment_Binary=attachment)
        self.session.add(new_submission)
        self.session.commit()
        log.info("Message from {} with eligible attached files committed to database.".format(submitter))
        i = new_submission.Id
        log.debug('Last message insert available on row %s in database' % str(i))
        return i

    # Deletes a message row from the database given its id
    def delete_message(self, db_id):
        message_to_delete = self.session.query(Submission).filter_by(Id=db_id).one()
        self.session.delete(message_to_delete)
        self.session.commit()
        log.info("Message on row id {} deleted")
        return True

    # Returns number of submissions as integer
    def countlines(self):
        line_count = self.session.query(Submission).count()
        log.debug("Found {} submissions in database".format(line_count))
        return line_count

    def get_max_id(self):
        max_id = self.session.query(func.max(Submission.Id))
        log.info("Identified {} as highest index in database".format(max_id))
        return max_id

    # Returns a list of all Ids in submissions database
    def get_all_ids(self):
        res = self.session.query(Submission.Id).all()
        ids = [tup[0] for tup in res]
        return ids

    # Inserts results from the ExcelParser analysis in the database. Returns the id of the row edited.
    def insert_analysis(self, db_id, **kwargs):
        submission = self.session.query(Submission).filter_by(Id=db_id).one()

        submission.Lead_Office = kwargs.get('lead_office')
        submission.P_Margin = kwargs.get('project_margin')
        submission.Tot_Fee = kwargs.get('total_fee')
        submission.Blended_Rate = kwargs.get('blended_hourly_rate')
        submission.Tot_Hours = kwargs.get('total_hours')
        submission.Hours_Mgr = kwargs.get('hours_by_role').get('Manager')
        submission.Hours_SPM = kwargs.get('hours_by_role').get('SPM')
        submission.Hours_PM = kwargs.get('hours_by_role').get('PM')
        submission.Hours_Cons = kwargs.get('hours_by_role').get('Cons')
        submission.Hours_Assoc = kwargs.get('hours_by_role').get('Assoc')
        submission.Method = kwargs.get('pricing_method')
        submission.Tool_Version = kwargs.get('tool_version')

        self.session.add(submission)
        self.session.commit()
        log.info('Quote analysis inserted and committed to database on database ID {}'.format(db_id))
        assert submission.Id == db_id
        return db_id

    # Closes a session
    def close(self):
        self.session.close()

    # Retrieves a file from the database and returns the filename for it
    def get_file_by_id(self, i):
        log.debug("Retrieving file with database index {}".format(i))
        submission = self.session.query(Submission).filter_by(Id=i).one()
        tempfile_name = "__Written_From_DB__.xlsm"
        tempfile_contents = submission.Attachment_Binary

        with open(tempfile_name, "wb") as f:
            f.write(tempfile_contents)
        return tempfile_name

    # Returns list with all values from a specified column - HARD DEPRECIATION
    def get_column_values(self, col):
        raise NotImplementedError

    # __init__ starts up a connection to a given database
    # Access credentials are taken from a "DBcreds.py" in a subdirectory "credentials"
    # DBcreds.py must include a variable "user" and one "password"
    def __init__(self):
        if config.database == 'MySQL on Azure':
            raise NotImplementedError

        elif config.database == 'Sqlite3':
            raise NotImplementedError

        elif config.database == 'MSSQL on Azure':
            self.engine = create_engine('mssql+pyodbc://{0}:{1}@{2}/{3}?driver=ODBC+Driver+13+for+SQL+Server'.format(DBcreds.user, DBcreds.password, DBcreds.host, DBcreds.database))
            Base.metadata.create_all(self.engine)
            Session = sessionmaker(bind=self.engine)
            self.session = Session()
            log.info("Setup done for MSSQL on Azure, using SQLAlchemy ORM")

        else:
            print("Error configuring database connection")
            log.error("Fatal error configuring database")

class DuplicateFileWarning(Exception):
    pass


class DuplicateMessageWarning(Exception):
    pass
