import getcalcs
import logging
from excel_parser import ExcelParsingError
from DBhelper import DBHelper

def main():
    db = DBHelper()
    db_rows = db.countlines()

    for row in range(1,db_rows+1):
        try:
            getcalcs.analyze_submission(row)
        except ExcelParsingError as error:
            logging.error(repr(error))


if __name__=='__main__':
    main()