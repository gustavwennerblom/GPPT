import logging
import sys
import unittest
from DBhelper_sqla import DBHelper
from MyMessage import MyMessage
from sqlalchemy import desc
import config
if config.test_database:
    from models import DevSubmission as Submission, DevLastUpdated as LastUpdated
else:
    from models import Submission, LastUpdated

class TestDBHelper(unittest.TestCase):
    def setUp(self):
        logging.basicConfig(stream=sys.stdout, format=config.LOG_FORMAT, level=config.loglevel)
        self.db = DBHelper()

    def test_set_timestamp_now(self):
        timestamp_set = self.db.set_timestamp()
        timestamp_got = self.db.get_timestamp()
        timestamp_db = self.db.session.query(LastUpdated).filter_by(Id=1).one().Updated
        self.assertEqual(timestamp_set, timestamp_got)
        self.assertEqual(timestamp_set, timestamp_db)

    def test_duplicate_file(self):
        filename_in_db = 'Copy of Copy of BuS Global Pricing tool v0.64 - Quant.xlsm 06-okt-17 15-03-51.xlsm'
        filename_not_in_db = 'DUMMY TITLE FROM GUSTAV'
        self.assertTrue(self.db.duplicate_file(filename=filename_in_db))
        self.assertFalse(self.db.duplicate_file(filename=filename_not_in_db))

    def test_duplicate_message(self):
        message_id_in_db = 'AAMkADBmYmFkNDU1LTE4MWEtNDdhNy05ODQxLTIwOGNlMWE0ZTJhYwBGAAAAAACuJ+e/J+cnSL+7PeCzc5F0BwBHzg3dfEohTpJuuhlo3bqdAAAAxZFtAAD1bJ37bup/Tb1FMrEI5mURAABwybayAAA='
        message_id_not_in_db = 'BLAJ'
        self.assertTrue(self.db.duplicatemessage(message_id=message_id_in_db))
        self.assertFalse(self.db.duplicatemessage(message_id=message_id_not_in_db))

    def test_line_count(self):
        line_count = self.db.countlines()
        self.assertIsInstance(line_count, int)
        self.assertGreater(line_count, 0)

    def test_insert_message(self):
        my_message = MyMessage()
        mess = my_message.get_message()
        insert_index = self.db.insert_message(mess.attachments[0].name,
                                              mess.sender.email_address,
                                              mess.subject,
                                              str(mess.datetime_received),
                                              str(mess.item_id),
                                              str(mess.attachments[0].attachment_id),
                                              mess.attachments[0].content)
        self.assertIsInstance(insert_index, int)
        self.test_insert_id = insert_index
        self.db.delete_message(db_id=self.test_insert_id)

    def test_insert_analysis(self):
        max_index = self.db.session.query(Submission.Id).order_by(desc(Submission.Id)).first()[0]
        insert_index = self.db.insert_analysis(db_id=max_index,
                                               lead_office = "TEST_OFFICE",
                                               project_margin = 99.99,
                                               total_fee = 99999,
                                               total_hours = 88888,
                                               hours_by_role = dict(Mgr=99, SPM=99, PM=99, Cons=99, Assoc=99),
                                               blended_hourly_rate = 98.76,
                                               pricing_method = "TEST_METHOD",
                                               tool_version = "TEST_VERSION")

        updated_submission = self.db.session.query(Submission).filter_by(Id=insert_index).one()
        self.assertEqual(updated_submission.Lead_Office, "TEST_OFFICE")
        self.assertEqual(updated_submission.Id, max_index)

    def tearDown(self):
        self.db.session.close()