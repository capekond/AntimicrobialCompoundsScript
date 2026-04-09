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

if __name__ == '__main__':
    unittest.main()