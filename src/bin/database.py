from typing import List

from pandas import DataFrame
from sqlalchemy import create_engine
from tabulate import tabulate
import sqlite3
from src.bin.excel_parser import ExcelParser

class Database(ExcelParser):
    def __init__(self):
        super().__init__()
        self.conn = sqlite3.connect(self.DATABASE)
    def __exit__(self):
        self.conn.close()

    def db_execute(self, sql):
        self.log.debug("Executing SQL: " + sql)
        return self.conn.execute(sql)

    def exists_type_essay(self) -> bool:
        c = self.db_execute(f"SELECT COUNT(*) FROM {self.TABLE_NAME} WHERE activity IN ('{"','".join(self.p.type_essay)}');")
        return c.fetchall()[0][0] != 0


    def write_data(self, df: DataFrame):
        engine = create_engine(f'sqlite:///{self.DATABASE}', echo=False)
        return df.to_sql(self.TABLE_NAME, con=engine, if_exists='append')

    def get_overview(self):
        tc = self.db_execute(f"SELECT COUNT(*) FROM  {self.TABLE_NAME};")
        c = self.db_execute(f"SELECT timestamp, COUNT(*) FROM  {self.TABLE_NAME} GROUP BY timestamp ORDER BY timestamp;")
        self.log.info(f" Record in database: {tc.fetchone()[0]} Database content overview:\n" + tabulate(c, headers=[description[0] for description in c.description], tablefmt='grid'))

    def delete_data(self):
        ts = self.expand_range_sql()
        sql = f"DELETE FROM {self.TABLE_NAME} " + (f" WHERE timestamp  {ts};" if ts else ";")
        self.db_execute(sql)
        self.conn.commit()
        self.log.info(" Delete records with timestamps " + (self.p.list or "ALL "))

    def join_data(self):
        ts_final = self.get_ts()
        ts = self.expand_range_sql()
        sql = f"UPDATE {self.TABLE_NAME} SET timestamp = '{ts_final}' " + (f" WHERE timestamp  {ts};" if ts else ";")
        self.db_execute(sql)
        self.conn.commit()
        self.log.info(" Join records from " + (ts or "ALL ") + f"to {ts_final}")

    def expand_range_sql(self):
        if self.p.range:
            sql = f"SELECT timestamp FROM script_data WHERE timestamp >= '{self.p.range[0]}' and timestamp <= '{self.p.range[1]}' GROUP BY timestamp"
            cf = self.db_execute(sql)
            if not cf.fetchall():
                rng = " - ".join(self.p.range)
                self.log.error(f" No records in range {rng} ")
                self.p.list = "Empty, no records"
                return " IN () "
            self.p.list = [t[0] for t in cf]
        return " IN ('" + "','".join(self.p.list) + "')" if self.p.list else None

    def check_wrong_essay(self, check_db: bool = False) -> bool:
        self.p.type_essay = self.p.type_essay or self.ACTIVITIES
        for ts in self.p.type_essay:
            if not ts in self.ACTIVITIES:
                self.log.error(f" Wrong value for type essay / activity: {ts}")
                return False
        if check_db:
            return self.exists_type_essay()
        return True
