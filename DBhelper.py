import sqlite3


class DBHelper:

    def insert(self, filename, submitter, region, date, attachment_id, attachment):
        conn = sqlite3.Connection("submissions.db")
        cur = conn.cursor()
        cur.execute(
            '''INSERT INTO Test_Submissions
            (Filename, Submitter, Region, Date, Attachment_Id, Attachment_Binary)
            VALUES (?, ?, ?, ?, ?, ?)''',
            (filename, submitter, region, date, attachment_id, attachment)
        )
        conn.commit()
        print("Insert committed to database.")
        conn.close()
        print("Database connection closed")

    def __init__(self, dbname="submissions.db"):
        self.dbname=dbname

