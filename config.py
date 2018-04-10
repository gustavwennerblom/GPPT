import logging
loglevel = logging.INFO
debug = False
enforce_unique_messages = True
enforce_unique_files = True
# Options: "MySQL on TestRig", "MySQL on Azure", "Sqlite3", "MSSQL on Azure"
database = "MSSQL on Azure"
test_database = True
LOG_FORMAT = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
