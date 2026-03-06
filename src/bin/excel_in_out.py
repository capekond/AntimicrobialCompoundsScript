import logging

import openpyxl
import pandas
from openpyxl.styles import Alignment, Font
import pandas as pd
from sqlalchemy import create_engine
from src.bin.database import Database

class ExcelInOut(Database):

    def __init__(self):
        super().__init__()
        ts0 = self.expand_range_sql()
        ts = " AND timestamp " + ts0 if ts0 else ""
        self.SQL_FINAL = """
                SELECT pathogen, code, item_value, count(item_value) as cnt_item 
                    FROM  """ + self.TABLE_NAME + """ 
                    WHERE activity = '{activity}' """ + ts + """
                      GROUP BY pathogen, code, activity, item_value
                HAVING count(item_value) > 1
            """

    def import_source(self):
        wbi = self.open_file(openpyxl.load_workbook, self.p.import_source)
        return self.write_data(self.get_db_data(wbi))

    def import_backup_excel(self):
        df = pd.read_excel(self.p.import_backup)
        engine = create_engine(f'sqlite:///{self.DATABASE}', echo=False)
        df.to_sql(self.TABLE_NAME, con=engine, if_exists='append', index=False)
        return len(df)

    def export_backup_excel(self):
        ts = self.expand_range_sql()
        sql = f"SELECT row_id, sheet, code, pathogen, activity, item, item_value, timestamp FROM script_data {'WHERE timestamp' + ts + ';' if ts else ';'}"
        df = pd.read_sql_query(sql, self.conn)
        df.to_excel(self.p.export_backup, index=False, engine='openpyxl')
        return len(df)

    def get_final_content(self):
        with pd.ExcelWriter(self.p.export_final, engine='openpyxl') as writer:
            for type_essay in self.p.type_essay:
                sql = self.SQL_FINAL.format(activity=type_essay)
                logging.debug("Executing SQL: " + sql)
                df = pandas.read_sql_query(sql, self.conn)
                pivot_df = df.pivot_table(values='item_value', index=['code'], columns=['pathogen'], aggfunc="first")
                pivot_df.to_excel(writer, sheet_name=type_essay, index=True)
                for c in writer.sheets[type_essay]['A1':'AA1'][0]:
                    c.alignment = Alignment(textRotation=90)
                    c.font = Font(bold=False)