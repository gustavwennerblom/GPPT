#NOTE - UNTESTED

import sqlite3
conn = sqlite3.Connection("submisisons.db")
cur = conn.cursor()
cur.execute("CREATE TABLE IF NOT EXISTS Test_Submissions (Id INTEGER PRIMARY KEY, Filename TEXT, Submitter TEXT, Region TEXT, Date TEXT, Attachment_Id TEXT, Attachment_Binary BLOB)")

