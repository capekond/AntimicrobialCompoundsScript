from typing import Callable, Tuple, List

from src.bin.excel_in_out import ExcelInOut
class Main(ExcelInOut):
    def __init__(self):
        super().__init__()

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
            if self.p.import_source:
                self.report_errors()
            self.get_overview()
        else:
            # for condition, operation in self._get_operations():
            #     if condition:
            #         operation()


            if self.p.delete:
                self.delete_data()
            if self.p.import_source:
                self.import_source()
            if self.p.join:
                self.join_data()
            if self.p.import_backup:
                self.import_backup_excel()
            if self.p.export_backup:
                self.export_backup_excel()
            if self.p.export_final:
                self.get_final_content()


if __name__ == "__main__":
    m = Main()
    m.main()
