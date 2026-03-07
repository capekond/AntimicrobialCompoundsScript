from typing import Callable, Tuple, List

from src.bin.excel_in_out import ExcelInOut
class Main(ExcelInOut):
    def __init__(self):
        super().__init__()

    def dry_run(self) -> None:
        if self.p.import_source:
            self.report_errors()
        self.get_overview()

    def _get_operations(self) -> List[Tuple[bool, Callable]]:
        """Return list of (condition, operation) tuples."""
        return [
            (self.p.delete, self._execute_delete),
            (self.p.import_source, self._execute_import_source),
            (self.p.join, self._execute_join),
            # ... etc
        ]

    def main(self) -> None:
        self.check_args()


        if self.p.dry_run:
            self.dry_run()
        else:
            if self.p.delete:
                self.delete_data()
                self.log.info(" Delete records with timestamps " + (self.p.list if self.p.list else "ALL "))
            if self.p.import_source:
                info = self.import_source()
                self.log.info(f" Update affect {info} rows in data table.")
            if self.p.join:
                ts, tss = self.join_data()
                self.log.info(" Join records from " + (tss if tss else "ALL ") + f"to {ts}")
            if self.p.import_backup:
                info = self.import_backup_excel()
                self.log.info(f" Import exported data from file {self.p.import_backup} to database data.db. Row count is {info}")
            if self.p.export_backup:
                info = self.export_backup_excel()
                self.log.info(f" Backup database data to {self.p.export_backup} Row count is {info}")
            if self.p.export_final:
                self.get_final_content()
                self.log.info(f" Scope {self.p.list if self.p.list else "ALL"} exported to final data {self.p.export_final}")
        self.conn.close()

if __name__ == "__main__":
    m = Main()
    m.main()
