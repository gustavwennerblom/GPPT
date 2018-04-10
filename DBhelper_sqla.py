import logging
import config
from datetime import datetime
import credentials.DBcreds as DBcreds
from sqlalchemy import create_engine, Table, String, Integer, MetaData
from sqlalchemy.orm import sessionmaker
from sqlalchemy.ext.declarative import declarative_base
from models import Base, Submission, LastUpdated

log = logging.getLogger(__name__)
log.setLevel(logging.DEBUG)
ch = logging.StreamHandler()
ch.setLevel(logging.DEBUG)
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
ch.setFormatter(formatter)
log.addHandler(ch)

class DBHelper:

    # Sets a timestamp (string) in the 'latest updated' table in the database and returns it (as string)
    def set_timestamp(self, *args):
        if len(args) > 2:
            logging.error('set_timestamp takes maximum one argument, %i given' % len(args))
            raise TypeError("set_timestamp in DBhelper misused. Check log.")
        elif len(args) == 1:
            timestamp = args[0]
        else:
            t = datetime.now()
            timestamp = datetime.strftime(t, "%Y-%m-%d-%H-%M-%S")

        res = self.session.query(LastUpdated).filter_by(Id=1).one()
        res.Updated = timestamp
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

    # Checks if a given message id is in the database. Returns boolean.
    def duplicatemessage(self, message_id):
        log.debug("Checking if message with EWS id {} is in already in database".format(message_id))
        hits = self.session.query(Submission).filter_by(Message_Id=message_id).all()
        return bool(hits)

    # Inserts an set of data as a new row in the database. Returns the index of that last insert.
    def insert_message(self, filename, submitter, region, date, message_id, attachment_id, attachment):
        raise NotImplementedError

    # Returns number of submissions as integer
    def countlines(self):
        line_count = self.session.query(Submission).count()
        log.debug("Found {} submissions in database".format(line_count))
        return line_count

    def insert_analysis(self, db_id, **kwargs):
        raise NotImplementedError

    # Closes a session
    def close(self):
        self.session.close()

    # Retrieves a file from the database and returns the filename for it
    def get_file_by_id(self, i):
        raise NotImplementedError

    # Returns list with all values from a specified column
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

        else:
            print("Error configuring database connection")

class DuplicateFileWarning(Exception):
    pass


class DuplicateMessageWarning(Exception):
    pass


if __name__ == '__main__':
    db = DBHelper()
    print(db.get_timestamp())
