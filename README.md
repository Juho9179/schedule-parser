# Schedule-parser

# What
Schedule-parser is a simple Python script that parses a workshift schedule which is in .xlsx format and has multiple employees in same file.

Schedule-parser then parses given employee's shifts and lists them in Google Calendar compatible format in CSV.
## WARNING: OVERWRITES old file, should it exist.

# How to run
```
python3 .\schedule-parser.py EMPLOYEE-NAME '.\schedule.xlsx'
```
Scheduler-parser then creates a new file, in the same directory as .xlsx file with same name, but .csv extension.

# Dependencies

openpyxl: https://openpyxl.readthedocs.io/en/stable/

"openpyxl is a Python library to read/write Excel 2010 xlsx/xlsm/xltx/xltm files."

How to install:

```
pip install openpyxl
```

# Example
dummy_data.xlsx
![alt text](/docs/material.png)

```
python3 .\schedule-parser.py AB18 '.\example\dummy_data.xlsx'
-> outputs AB18's shifts in .\example\dummy_data.xlsx.csv
```

dummy_data.xlsx.csv

![alt text](/docs/result.png)

Scheduler-parser then creates a new file, in the same directory as .xlsx file with same name, but .csv extension.