# AntimicrobialCompounds
The increase in the number of bacterial strains resistant to known antibiotics combined with decrease of new antibiotic introduced to clinic is alarming. Our group is attempting to contribute to solve this serious problem

## Prerequisites 
- python 3.12
# SCRIPT
```
options:

  -h, --help            show this help message and exit

Generic arguments:
  -v, --verbose         Verbose output
  -n, --no_question     Disable approval question
  -D, --dry_run         Dry run: If present -I validate Excel import data. Optional argument is file name with dry run results.
                        If present -R provide info of raw data file (counts group counts by timestamps), use with -v

Input data:

  -i, --import [Excel file name]
                        Import Excel file with data sources to database. To check the content use parameter --dry_run.
  -s, --sheets [list of sheets in import file]
                        Source worksheets. If missing all worksheets will be used in given import file or database

Import data to database:

  -r, --is_range [FROM TO]
                        -d -j -g use 2 value as boundaries form ... to. The rounding i.e 2025.12 could be used
  -l, --is_list  [names of list]
                        -d -j -g use values as list.
  -d, --delete
                        Delete records with timestamps as the list or range (if -r is present)
  -j, --join 
                        Join records with timestamps as the list or range (if -r is present). Actual timestamp as is used for joined data.

Final data file manipulation:

  -e, --export Exported [final excel file]
                        Default value example: C:\Users\ocape\PycharmProjects\AntimicrobialCompounds\src\script\export-1767888617.xlsx

  -t, --type_essay [TYPE_ESSAY ...] 
                        MIC and / or MBC, sheets in export file 
                        
```

# GUI
## Used tools 
https://nicegui.io/

https://www.sqlite.org/

## Limitations
- no deployment script   
- no password, no encryption database 
- use only at intranet
- no performance optimization
## TODO
5. catch and log errors, provide info to user
7. requirements file
8. fix problem with update on  main page
10. export pdf
## DONE
1. user update database and refresh table
6. use property file for global variable
9. disable ADD RECORD
10. user access login / logout
4. log folder in git
11. links to shared function
12. enable and disable login
13. can change welcome screen by markdown