# SQL-statement-generator-from-EXCEL-file


# Description
A python script that analyzes an excel file and generates a SQL statement to an existing DB.  
The script should be called from a CLI with `EXCEL` file name as argument.

# ScreenShots
## INPUT: EXCEL file
<img src="https://user-images.githubusercontent.com/58184521/122626269-4082c480-d0b2-11eb-89b0-c7b979c1637f.png" width="650">

  ## OUTPUT: SQL statement
<img src="https://user-images.githubusercontent.com/58184521/122626314-7aec6180-d0b2-11eb-9927-a44fa60629b2.png" width="650">

# Script's assumptions
- The `EXCEL` will be in the same directory as the script file.
- The `EXCEL` file name, will be given to the script as argument, will be given without extension.
- The `EXCEL` file will follow the script's 'EXCEL' file conventions.

# 'EXCEL' file conventions
- The file should have ONE sheet only.
- The first row will be the columns titles.
- The columns titles will match the existing DB columns titles.
- The rows will start at [0,0] cell (top left).
- All the values types will match the DB columns types.

# PK column
- The `email` column is set as PK in the existing DB therefore duplicate `email` values are not allowed.
- The script will check for double `email` values, in case of double, the script will notify and abort.

# Requirements

1.  [Python 3.9.3 (or above).](https://www.python.org/downloads/)
2.  [Microsoft Excel](https://www.microsoft.com/en-us/microsoft-365/excel).

# Installation & Run
 
1. Clone [SQL-statement-generator-from-EXCEL-file](https://github.com/AmitAharoni/SQL-statement-generator-from-EXCEL-file) repository.
   ```sh
   https://github.com/AmitAharoni/SQL-statement-generator-from-EXCEL-file.git
   ```
2. Open CLI (Command Line Interface).
3. Navigate to the project directory.
4. Place a valid `EXCEL` file in the directory.
5. run the command
     ```sh
   python SQL_statement_generator_from_EXCEL_file.py <EXCEL FILE NAME>
   ```
