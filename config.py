# Logging settings
import logging
loglevel = logging.INFO
LOG_FORMAT = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
log_email_recipient = "gustav.wennerblom@gmail.com"
log_directory = "logs"
main_log_filename = "main.log"

# Development vs production settings
debug = False
test_database = True

# General settings
enforce_unique_messages = True
enforce_unique_files = True
database = "MSSQL on Azure"     # Options: "MySQL on TestRig", "MySQL on Azure", "Sqlite3", "MSSQL on Azure"
