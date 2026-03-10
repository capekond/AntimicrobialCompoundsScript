from typing import Callable, Tuple, List

from src.bin.excel_in_out import ExcelInOut


class Main(ExcelInOut):
    def __init__(self):
        super().__init__()

    def _dry_run(self):
        if self.p.import_source:
            self.report_errors()
        self.get_overview()

    def _get_operations(self) -> List[Tuple[bool, Callable]]:
        return [
            (self.p.dry_run, self._dry_run),
            (self.p.delete, self.delete_data),
            (self.p.import_source, self.import_source),
            (self.p.join, self.join_data),
            (self.p.import_backup, self.import_backup_excel),
            (self.p.export_backup, self.export_backup_excel),
            (self.p.export_final, self.get_final_content)
        ]

    def main(self) -> None:
        self.check_args()
        for condition, operation in self._get_operations():
            if condition:
                    operation()
                    if operation == self._dry_run:
                        exit(0)


if __name__ == "__main__":
    m = Main()
    m.main()
