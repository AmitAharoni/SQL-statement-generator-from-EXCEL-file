<img src="https://user-images.githubusercontent.com/58184521/122685801-632dee00-d216-11eb-8c34-4f57eac2dc20.png" width="650">

# Description
 A python script that analyzes an excel file and generates a SQL statement for existing DB with multiple tables.  

# Script arguments
The script should be called from a CLI interface with three arguments.
1. `EXCEL` file path. 
2. Club id.
3. PK field title.

# PK field column validation
- One of [first_name, last_name, email,	phone] will be set as PK in the existing DB.
- The PK column must exist in the file.
- Duplicate PK values are not allowed in the file.
- In any case of violation, the script will notify and abort.

# 'EXCEL' file conventions
- The file should have ONE sheet only.
- The first row will be the columns titles.
- The columns titles will inclued entire pre-known list, in any order.  
   [first_name, last_name,	email,	phone,	membershp_start_date,	membership_end_date,	membership_name].
- The rows will start at ['A1'] cell (top left).
- All the values types will match the DB columns types.

# ScreenShots
## INPUT: EXCEL file
<img src="https://user-images.githubusercontent.com/58184521/122684862-b8670100-d210-11eb-98b2-ba824c81b68f.png" width="650">

 ## OUTPUT: SQL statement
<img src="https://user-images.githubusercontent.com/58184521/122684876-d3397580-d210-11eb-9460-1a3d08c6a978.png" width="650">

## OUTPUT: users table
<img src="https://user-images.githubusercontent.com/58184521/122685782-2a8e1480-d216-11eb-9bc4-8356d637d058.png" width="650">

## OUTPUT: memberships table
<img src="https://user-images.githubusercontent.com/58184521/122685725-e1d65b80-d215-11eb-9da6-67d9f233504c.png" width="650">

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
4. run the command
     ```sh
   python SQL_statement_generator_from_EXCEL_file.py <EXCEL FILE PATH> <CLUB ID> <PK FIELD TITLE>
   ```
