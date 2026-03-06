import argparse
import datetime
import sys
from tabulate import tabulate
import logging

def log_for_level(self, message, *args, **kwargs):
    if self.isEnabledFor(45):
        self._log(45, message, args, **kwargs)

def log_to_root(message, *args, **kwargs):
    logging.log(45, message, *args, **kwargs)

class Arguments:

    def __init__(self):
        ts = int((datetime.datetime.now()).timestamp())
        self.SHOW_LOG_LEVEL = 45
        self.DATABASE = "data.db"
        self.TABLE_NAME = 'script_data'
        self.ACTIVITIES = ("MIC", "MBC")
        self.ITEMS = (1, 2, 3, 1, 2, 3)
        self.COLUMNS = ("sheet", "row_id", "code", "pathogen", "activity", "item", "item_value", "timestamp")
        self.COLUMNS_ERR = ("sheet", "cell", "actual value", "error_description")
        self.COLUMNS_INFO = ("sheet", "status" , "row_count", "err_count")
        self.ITEM_COL_OFFSET = 2
        self.CODE_REGEX = r"^\s*\d+\s*-[\s\S]{7,}$"
        self.p = self.get_args()
        logging.addLevelName(self.SHOW_LOG_LEVEL, 'SHOW')
        logging.basicConfig(format='%(asctime)s %(levelname)s :%(message)s', level=logging.DEBUG if self.p.verbose else self.SHOW_LOG_LEVEL)
        setattr(logging, "SHOW", self.SHOW_LOG_LEVEL)
        setattr(logging.getLoggerClass(), 'show', log_for_level)
        setattr(logging, 'show', log_to_root)
        self.log = logging.getLogger(__name__)

    def get_args(self):
        parser = argparse.ArgumentParser()
        generic = parser.add_argument_group('Generic arguments')
        imp = parser.add_argument_group('Import and export from / to files')
        db = parser.add_argument_group('Manipulation with database content, set scope for output')
        generic.add_argument("-v", "--verbose", action='store_true', help="Verbose output")
        generic.add_argument("-n", "--no_question", action='store_true', help="Disable approval question")
        generic.add_argument("-D", "--dry_run", action='store_true', help="Dry run: If present -I validate Excel import data. Optional argument is file name with dry run results. use with -v")
        generic.add_argument("-r", "--range", nargs=2, help="For exports and database updates use 2 timestamps as boundaries form ... to. The rounding i.e  2025.12 could be used")
        generic.add_argument("-l", "--list", nargs='+', help="For exports and database updates provide timestamps list .")
        imp.add_argument("-i", "--import_source", type=str, help=f"Import Excel file with data sources to database {self.DATABASE}. To check the content use parameter --dry_run. ")
        imp.add_argument("-s", "--sheets", nargs='+', type=str, help="Source worksheets. If missing all worksheets will be used in  given import file or database")
        imp.add_argument("-I", "--import_backup", type=str, help=f"Import exported data from file to database {self.DATABASE}. To check the content use parameter --dry_run. ")
        imp.add_argument("-E", "--export_backup", type=str, help=f"Import exported data from file to database {self.DATABASE}. To check the content use parameter --dry_run. ")
        imp.add_argument("-e", "--export_final", help=f"Export to final excel file.")
        imp.add_argument("-t", "--type_essay", nargs='+', help="MIC and / or MBC, sheets in final export file ")
        db.add_argument("-d", "--delete", action='store_true', help="Delete records with timestamps by range or list")
        db.add_argument("-j", "--join", action='store_true', help="Join records with timestamps as the list or range. Actual timestamp as is used for joined data")
        return parser.parse_args()

    def check_args(self):
        if len(sys.argv) == 1:
            print("Hi, no expected arguments, let try --help for beginning")
            exit(0)
        if not self.p.type_essay:
            self.p.type_essay = self.ACTIVITIES
        self.log.info("List of variables:\n" + tabulate((dict(vars(self.p))).items(), headers=["Variable", "Value"],tablefmt="grid"))
        if (not self.p.no_question) and (not input("Do you like to proceed the task? [Y/n]") == "Y"):
            self.log.info("Script terminated by user.")
            exit(0)

    def open_file(self, open_function, file_path: str):
        try:
            wbi = open_function(file_path)
        except (FileNotFoundError, PermissionError) as e:
            self.log.error(e)
            exit(1)
        self.p.sheets = self.p.sheets if self.p.sheets else wbi.sheetnames
        return wbi
