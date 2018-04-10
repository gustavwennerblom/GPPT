import unittest
from DBhelper_sqla import DBHelper
from models import LastUpdated

class TestDBHelper(unittest.TestCase):
    def setUp(self):
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

    def tearDown(self):
        self.db.session.close()