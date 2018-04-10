import logging
import config
from datetime import datetime
import credentials.DBcreds as DBcreds
from sqlalchemy import create_engine, Table, String, Integer, MetaData
from sqlalchemy.ext.declarative import declarative_base


class DBHelper:

    # Checks if a certain filename is already in the database
    # Valid since all files have a timestamp in the filename

    def set_timestamp(self, *args):
        if len(args) > 2:
            logging.error('set_timestamp takes maximum one argument, %i given' % len(args))
            raise TypeError("set_timestamp in DBhelper misused. Check log.")
        elif len(args) == 1:
            timestamp = args[0]
        else:
            t = datetime.now()
            timestamp = datetime.strftime(t, "%Y-%m-%d-%H-%M-%S")

        sql = "UPDATE submissions.Test_Last_Update SET Updated=? WHERE ID=1"
        result = self.cur.execute(sql, (timestamp,))
        self.conn.commit()
        logging.info('Database timestamp for latest update set to %s' % timestamp)

    def get_timestamp(self):
        sql = "SELECT (Updated) FROM submissions.Test_Last_UPDATE WHERE ID=1"
        return self.cur.execute(sql).fetchone()[0]

    def duplicate_file(self, filename):
        sql = "SELECT Id FROM submissions.GPPT_Submissions WHERE Filename = ?"
        try:
            duplicate_list = self.cur.execute(sql, (filename,)).fetchall()
            print("Variable duplicate_list is: {}".format(duplicate_list))
            if duplicate_list:
                return True
            else:
                return False
        except DatabaseError as err:
            logging.error("Error on duplicate check for file {} ".format(filename) + repr(err))

    def duplicatemessage(self, message):
        sql = "SELECT Id FROM submissions.GPPT_Submissions WHERE Message_Id = ?"
        self.cur.execute(sql, (message.item_id, ))
        if not self.cur.fetchone():
            return False
        else:
            return True

            # Inserts an set of data as a new row in the database. Returns the index of that last insert.
    def insert_message(self, filename, submitter, region, date, message_id, attachment_id, attachment):

        # Check if file is already in database and raise error if so
        if self.duplicate_file(filename) and config.enforce_unique_files:
            logging.warning('Attempt to insert file %s in database disallowed. Set enforce_unique_files in '
                            'config-py to "False" to allow' % filename)
            raise DuplicateFileWarning("File is already in database")

        self.cur.execute(
            '''INSERT INTO submissions.GPPT_Submissions
            (Filename, Submitter, Region, Date, Message_Id, Attachment_Id, Attachment_Binary)
            VALUES (?, ?, ?, ?, ?, ?, ?)''',
            (filename, submitter, region, date, message_id, attachment_id, attachment))
        i = self.cur.execute('SELECT SCOPE_IDENTITY()').fetchone()
        self.conn.commit()
        # logging.info("Message metadata and eligible attached files committed to database.")
        logging.info('Last message insert available on row %s in database' % str(i))
        return i

    def countlines(self):
        # type: () -> int
        sql = "SELECT COUNT(*) FROM submissions.GPPT_Submissions"
        self.cur.execute(sql)
        result = self.cur.fetchone()[0]

        return result

    def insert_analysis(self, db_id, **kwargs):
        self.cur.execute(
            '''UPDATE submissions.GPPT_Submissions SET
            Lead_Office=(?),
            P_Margin=(?),
            Tot_Fee=(?),
            Blended_Rate=(?),
            Tot_Hours=(?),
            Hours_Mgr=(?),
            Hours_SPM=(?),
            Hours_PM=(?),
            Hours_Cons=(?),
            Hours_Assoc=(?),
            Method=(?),
            Tool_Version=(?)
            WHERE ID = (?)''',
            (kwargs["lead_office"], kwargs["project_margin"], kwargs["total_fee"], kwargs["blended_hourly_rate"],
             kwargs["total_hours"], kwargs["hours_by_role"]["Manager"], kwargs["hours_by_role"]["SPM"],
             kwargs["hours_by_role"]["PM"], kwargs["hours_by_role"]["Cons"], kwargs["hours_by_role"]["Assoc"],
             kwargs["pricing_method"], kwargs["tool_version"], db_id)
        )

        self.conn.commit()
        logging.info('Quote analysis inserted and committed to database on database ID %s' % str(db_id))

    def close(self):
        self.conn.close()
        logging.info("Database connection closed")

    # Retrieves a file from the database and returns the filename for it
    def get_file_by_id(self, i):
        logging.info("Retrieving file with database index %s" % i)

        self.cur.execute("SELECT (Attachment_Binary) FROM submissions.GPPT_Submissions WHERE (ID=?)", (i,))
        row = self.cur.fetchone()
        tempfile_name = "__Written_From_DB__.xlsm"
        tempfile_contents = row[0]

        with open(tempfile_name, "wb") as f:
            f.write(tempfile_contents)
        return tempfile_name

    # Returns list with all values from a specified column
    def get_column_values(self, col):
        result = []
        sql = "SELECT ? FROM submissions.GPPT_Submissions;"
        out = self.cur.execute(sql, (col,)).fetchall()
        for item in out:
            result.append(item)
        return result

    # __init__ starts up a connection to a given database
    # Access credentials are taken from a "DBcreds.py" in a subdirectory "credentials"
    # DBcreds.py must include a variable "user" and one "password"
    def __init__(self):
        if config.database == 'MySQL on Azure':
            from mysql.connector import DatabaseError

            ## Parameters required if running MySQL
            user = DBcreds.user
            password = DBcreds.password
            host = DBcreds.host
            database = DBcreds.database

            ### mysql.connector connection string
            import mysql.connector
            self.conn = mysql.connector.connect(user=user,
                                                password=password,
                                                host=host,
                                                database=database)
            # buffering kwarg - see https://stackoverflow.com/questions/29772337/python-mysql-connector-unread-result-found-when-using-fetchone
            self.cur = self.conn.cursor(buffered=True)
            logging.info('Connection created to database "{0}" on {1}'.format(database, host))

        elif config.database == 'Sqlite3':
            ### SQLite connection string
            import sqlite3
            self.conn = sqlite3.Connection("submissions.db")
            self.cur = self.conn.cursor()

        elif config.database == 'MSSQL on Azure':
            import pyodbc
            # self.conn = pyodbc.connect('Driver={ODBC Driver 13 for SQL Server};'
            #                            'Server=tcp:phoenixdb01.database.windows.net,1433;'
            #                            'Database=gppt-submissions;'
            #                            'Uid=;'
            #                            'Pwd=;'
            #                            'Encrypt=yes;'
            #                            'TrustServerCertificate=no;'
            #                            'Connection Timeout=30;')
            connstring = 'Driver={0};Server={1};Database={2};Uid={3};Pwd={4};Encrypt=yes;TrustServerCertificate=no;Connection Timeout=30;'
            self.conn = pyodbc.connect(connstring.format('{ODBC Driver 13 for SQL Server}',
                                                         DBcreds.host,
                                                         DBcreds.database,
                                                         DBcreds.user,
                                                         DBcreds.password))
            self.cur = self.conn.cursor()

            self.engine = create_engine('mssql+pyodbc://{0}:{1}@{2}/{3}?driver=ODBC+Driver+13+for+SQL+Server'.format(DBcreds.user, DBcreds.password, DBcreds.host, DBcreds.database))
            Base = declarative_base()
            meta = MetaData()
            meta.bind = self.engine
            meta.reflect()
            self.submissions = Table('submissions.gppt_submissions', meta, autoload=True)
            self.last_update = Table('submissions.test_last_update', meta, autoload=True)



        else:
            print("Error configuring database connection")

class DuplicateFileWarning(Exception):
    pass


class DuplicateMessageWarning(Exception):
    pass
