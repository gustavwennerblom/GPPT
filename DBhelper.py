import sqlite3


class DBHelper:

    def insert(self, filename, submitter, region, date, attachment_id, attachment):
        self.cur.execute(
            '''INSERT INTO Test_Submissions
            (Filename, Submitter, Region, Date, Attachment_Id, Attachment_Binary)
            VALUES (?, ?, ?, ?, ?, ?)''',
            (filename, submitter, region, date, attachment_id, sqlite3.Binary(attachment))
        )
        self.conn.commit()
        print("Insert committed to database.")

    def commit_and_close(self):
        pass

    def close(self):
        self.conn.close()
        print("Database connection closed")

    #retrieves a file from the database and returns the filename for it
    def get_file_by_id(self,id):
        row=self.cur.execute("SELECT (Attachment_Binary) FROM Test_Submissions WHERE (ID=?)", (id,)).fetchone()
        tempfile_name="Written_From_DB.xlsm"
        tempfile_contents = row[0]
        f=open(tempfile_name, "wb")
        f.write(tempfile_contents)
        f.close()
        return tempfile_name


    def __init__(self, dbname="submissions.db"):
        self.dbname=dbname
        self.conn = sqlite3.Connection("submissions.db")
        self.cur = self.conn.cursor()
