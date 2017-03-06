#NOTE - UNTESTED

import sqlite3
conn = sqlite3.Connection("submissions.db")
cur = conn.cursor()
cur.execute('''CREATE TABLE IF NOT EXISTS Test_SubmissionsB (
    Id INTEGER PRIMARY KEY,
    Filename TEXT,
    Submitter TEXT,
    Region TEXT,
    Date TEXT,
    Attachment_Id TEXT,
    Attachment_Binary BLOB,
    Lead_Office TEXT,
    P_Margin REAL,
    Tot_Fee INTEGER,
    Blended_Rate REAL,
    Tot_Hours INTEGER,
    Hours_Mgr INTEGER,
    Hours_SPM INTEGER,
    Hours_PM INTEGER,
    Hours_Cons INTEGER,
    Hours_Assoc INTEGER,
    Method TEXT)
    ''')

cur.execute('''CREATE TABLE IF NOT EXISTS Test_Last_Update (
    Id INTEGER PRIMARY KEY,
    Updated TEXT)
    ''')