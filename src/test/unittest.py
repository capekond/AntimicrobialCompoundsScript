import sqlite3
import unittest

from bin import database

class TestStringMethods(unittest.TestCase):

    def test_database_error(self):
        db = database.Database()
        with self.assertRaises(Exception) as context:
            db.db_execute("WRONG SQL")
        self.assertEqual(context.exception.args[0], 'near "WRONG": syntax error')
        self.assertEqual(type(context.exception), sqlite3.OperationalError)

    def test_activity_exists(self):
        db = database.Database()
        db.p.type_essay = ["MIC"]
        self.assertTrue(db.exists_type_essay())
        self.assertTrue(db.check_wrong_essay())
        self.assertTrue(db.check_wrong_essay(True))

    def test_activity_not_exists(self):
        db = database.Database()
        db.p.type_essay = ["XXX"]
        self.assertFalse(db.exists_type_essay())
        self.assertFalse(db.exists_type_essay())
        self.assertFalse(db.check_wrong_essay(True))

if __name__ == '__main__':
    unittest.main()