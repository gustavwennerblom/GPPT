import sqlite3
from credentials import DBcreds
# import mysql.connector
from datetime import datetime
import config
import sys

if config.database=="Sqlite3":
    conn = sqlite3.Connection("submissions.db")

    cur = conn.cursor()

    usr_in = input('WARNING - This will delete all data in database. Type "yes" to proceed \n>>')
    if usr_in != "yes":
        print("Exiting script. Bye.")
        sys.exit(1)

    cur.execute("DROP TABLE IF EXISTS GPPT_Submissions")
    cur.execute("DROP TABLE IF EXISTS Test_Last_Update")

    cur.execute('''CREATE TABLE IF NOT EXISTS GPPT_Submissions (
            Id INTEGER PRIMARY KEY,
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

else:
    conn = mysql.connector.connect(user=DBcreds.user,
                                   password=DBcreds.password,
                                   host=DBcreds.host,
                                   database=DBcreds.database)


    cur = conn.cursor()

    cur.execute('''CREATE TABLE IF NOT EXISTS GPPT_Submissions (
        Id INTEGER PRIMARY KEY AUTO_INCREMENT,
        Filename TEXT,
        Submitter TEXT,
        Region TEXT,
        Date TEXT,
        Message_Id TEXT,
        Attachment_Id TEXT,
        Attachment_Binary MEDIUMBLOB,
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
        Id INTEGER PRIMARY KEY AUTO_INCREMENT,
        Updated TEXT)
        ''')

t = datetime.now()
timestamp = datetime.strftime(t, "%Y-%m-%d-%H-%M-%S")

cur.execute("INSERT INTO Test_Last_Update (Updated) VALUES (?);", (timestamp,))
conn.commit()

print("Fresh database all set up")

