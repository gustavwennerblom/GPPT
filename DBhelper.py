import config
import sqlite3
import logging
from datetime import datetime


class DBHelper:

    # Checks if a certain filename is already in the database
    # Valid since all files have a timestamp in the filename

    def set_timestamp(self, *args):
        if len(args) > 2:
            log.error('set_timestamp takes maximum one argument, %i given' % len(args))
            raise TypeError("set_timestamp in DBhelper misused. Check log.")
            return None
        elif len(args) == 1:
            timestamp = args[0]
        else:
            t = datetime.now()
            timestamp = datetime.strftime(t, "%Y-%m-%d-%H-%M-%S")

        sql = "UPDATE Test_Last_Update SET Updated=? WHERE ID=1"
        self.cur.execute(sql, (timestamp,))
        self.conn.commit()
        logging.info('Database timestamp for latest update set to %s' % timestamp)

    # Tagged for removal - duplicate with set_timestamp
    # def update_timestamp(self):
    #     timestamp = datetime.strftime(datetime.now(), "%Y-%m-%d-%H-%M-%S")
    #     sql = "UPDATE Test_Last_Update SET Updated=? WHERE ID=1"
    #     self.cur.execute(sql, (timestamp,))
    #     self.conn.commit()

    def get_timestamp(self):
        sql="SELECT (Updated) FROM Test_Last_UPDATE WHERE ID=1"
        return self.cur.execute(sql).fetchone()

    def duplicate_file(self, filename):
        sql = "SELECT Id FROM Test_SubmissionsC WHERE Filename = ?"
        self.cur.execute(sql, (filename,))
        if not self.cur.fetchone():
            return False
        else:
            return True

    def duplicatemessage(self, message):
        sql = "SELECT Id FROM Test_SubmissionsC WHERE Message_Id = ?"
        self.cur.execute(sql, (message.item_id, ))
        if not self.cur.fetchone():
            return False
        else:
            return True

            # Inserts an set of data as a new row in the database. Returns the index of that last insert.
    def insert_message(self, filename, submitter, region, date, message_id, attachment_id, attachment):
        if self.duplicate_file(filename) and config.enforce_unique_files:
            logging.warning('Attempt to insert file {0} in database disallowed. Set enforce_unique_files in '
                            'config-py to "False" to allow'.format(filename))
            raise DuplicateFileError("File is already in database")
            return None

        self.cur.execute(
            '''INSERT INTO Test_SubmissionsC
            (Filename, Submitter, Region, Date, Message_Id, Attachment_Id, Attachment_Binary)
            VALUES (?, ?, ?, ?, ?, ?, ?)''',
            (filename, submitter, region, date, message_id, attachment_id, sqlite3.Binary(attachment))
        )
        self.conn.commit()
        logging.info("Message metadata and eligible attached files committed to database.")

        i = self.cur.execute("SELECT last_insert_rowid()").fetchone()[0]
        logging.info('Last message insert available on row %s in database' % str(i))
        return i

        # Perhaps a conn.close is needed here?

    def insert_analysis(self, db_id, **kwargs):
        self.cur.execute(
            '''UPDATE Test_SubmissionsC SET
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
            Method=(?)
            WHERE ID = (?)''',
            (kwargs["lead_office"], kwargs["project_margin"], kwargs["total_fee"], kwargs["blended_hourly_rate"],
             kwargs["total_hours"], kwargs["hours_by_role"]["Manager"], kwargs["hours_by_role"]["SPM"],
             kwargs["hours_by_role"]["PM"], kwargs["hours_by_role"]["Cons"], kwargs["hours_by_role"]["Assoc"],
             kwargs["pricing_method"], db_id)
        )

        self.conn.commit()
        logging.info ('Quote analysis inserted and committed to database on database ID %s' % str(db_id))

        # Perhaps a conn.close is needed here?

    def close(self):
        self.conn.close()
        logging.info("Database connection closed")

    # Retrieves a file from the database and returns the filename for it
    def get_file_by_id(self, i):

        logging.debug("Retrieving file with database index %s" % i)
        row = self.cur.execute("SELECT (Attachment_Binary) FROM Test_SubmissionsC WHERE (ID=?)", (i,)).fetchone()
        tempfile_name = "Written_From_DB.xlsm"
        tempfile_contents = row[0]
        f = open(tempfile_name, "wb")
        f.write(tempfile_contents)
        f.close()
        return tempfile_name

    def __init__(self, dbname="submissions.db"):
        self.dbname = dbname
        self.conn = sqlite3.Connection("submissions.db")
        self.cur = self.conn.cursor()

class DuplicateMessageError(Exception):
    pass

class DuplicateFileError(Exception):
    pass