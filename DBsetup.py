

# import sqlite3
# conn = sqlite3.Connection("submissions.db")
from credentials import DBcreds
import mysql.connector

conn = mysql.connector.connect(user=DBcreds.user,
                               password=DBcreds.password,
                               host='127.0.0.1',
                               database="submissions")
cur = conn.cursor()

cur.execute('''CREATE TABLE IF NOT EXISTS GPPT_Submissions (
    Id INTEGER PRIMARY KEY AUTO_INCREMENT,
    Filename TEXT,
    Submitter TEXT,
    Region TEXT,
    Date TEXT,
    Message_Id TEXT,
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
    Method TEXT,
    Tool_Version TEXT,
    SaleID INTEGER,
    ProjNo INTEGER)
    ''')

cur.execute('''CREATE TABLE IF NOT EXISTS Test_Last_Update (
    Id INTEGER PRIMARY KEY,
    Updated TEXT)
    ''')