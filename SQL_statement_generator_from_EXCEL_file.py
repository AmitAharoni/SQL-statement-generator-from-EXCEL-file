from datetime import datetime
import sys
import pathlib
import openpyxl
from openpyxl.utils import get_column_letter


def main(excel_file_name):
##  The main method of the script.
##  The script should be called from a CLI with `EXCEL` file name as argument.
##  The script generates an SQL statement from an excel file to an existing DB.
##  Script assumptions:
##  - The `EXCEL` file must be in the same directory as the script file.
##  - The `EXCEL` file name, given to the script as argument should be without extension.
##  - The `EXCEL` file should follow the company convention:
##    * The file should have ONE sheet only.
##    * The first row will be the columns titles.
##    * The columns titles will match the DB columns titles.
##    * The rows will start at [0,0] cell (top left).
##    * Rest of the file rows will have appropriate type values.
##    * The `email` column is set as PK in the company DB therefore duplicate `email` values are not allowed.
##      The script will check for double `email` values, in case of double, the script will abort.

    dir_path = pathlib.Path(__file__).parent.absolute()
    slash = '\\'
    extension = ".xlsx"
    path = dir_path.__str__() + slash + excel_file_name + extension

    try:
        workbook = openpyxl.load_workbook(path)
    except FileNotFoundError:
        print(f"The file `{excel_file_name}{extension}` wasn't found at `{dir_path}{slash}`.\n\
Please follow the instructions:.\n\
*The file must be in the same directory as the script file.\n\
*The file name, given to the script should be without extension.\n\
aborting script!")
        sys.exit()

    sheet = workbook.worksheets[0]
    ##pk_double_validation(sheet, "email")

    row_values_dictonary = {}
    num_of_cols = sheet.max_column
    num_of_rows = sheet.max_row
    users_sql_statement = "INSERT INTO users (first_name, last_name, phone, email, joined_at, club_id) \nVALUES \n"
    memberships_sql_statement = "INSERT INTO memberships (start_date, end_date, membership_name) \nVALUES \n"

    for i in range (2, num_of_rows + 1):
        for j in range (1, num_of_cols + 1):
            column_title = sheet.cell(row = 1, column = j).value
            cell_value = sheet.cell(row = i, column = j).value

            if type(cell_value) == datetime:
                row_values_dictonary[column_title] = str(cell_value.date())
            else:
                if cell_value is None:
                    row_values_dictonary[column_title] = "NULL"
                else:    
                    row_values_dictonary[column_title] = str(cell_value) 

        users_sql_statement = add_row_to_users_statement(users_sql_statement, row_values_dictonary)
        memberships_sql_statement = add_row_to_memberships_statement(memberships_sql_statement, row_values_dictonary)

    users_sql_statement = correct_sql_statement(users_sql_statement)
    memberships_sql_statement = correct_sql_statement(memberships_sql_statement)

    print(users_sql_statement)
    print()
    print(memberships_sql_statement)


def pk_double_validation(sheet, pk_title):
##  The method gets sheet and PK title.
##  If no such PK column or there is a double PK value, abort the script.

    pk_dictonary = {}
    pk_col_as_int = get_title_col(sheet, pk_title)
    pk_col_as_letter = get_column_letter(pk_col_as_int)

    if pk_col_as_int == -1:
        print(f"No `{pk_title}` column  in the file!\naborting script")
        sys.exit()
    for pk in sheet[pk_col_as_letter]:
        if pk_dictonary.__contains__(pk.value) is False:
            pk_dictonary[pk.value] = pk.value
        else:
            print(f"The PK - `{pk.value}`, exists DOUBLE or more in the file.\naborting script")
            sys.exit()

def get_title_col(sheet, title):
## The method gets sheet and title.
## Returns the index of the title's col, if no such a title, returns -1.

    num_of_cols = sheet.max_column
    for i in range (1, num_of_cols + 1):
        if sheet.cell(row = 1, column = i).value == title:
            return i
    return -1

def add_row_to_users_statement(users_sql_statement, row_values_dictonary):
## The method gets 'users' SQL statement and row values.
## Returns a combined SQL statement of the statement and the values.

    users_row = (f"      (\
{row_values_dictonary['first_name']}, {row_values_dictonary['last_name']},\
 {row_values_dictonary['phone']}, {row_values_dictonary['email']},\
 {row_values_dictonary['membershp_start_date']}, 2400),\n")
    users_sql_statement += users_row
    return users_sql_statement

def add_row_to_memberships_statement(memberships_sql_statement, row_values_dictonary):
## The method gets 'memberships' SQL statement and row values.
## Returns a combined SQL statement of the statement and the values.

    memberships_row = (f"      (\
{row_values_dictonary['membershp_start_date']}, {row_values_dictonary['membership_end_date']},\
 {row_values_dictonary['membership_name']}),\n")
    memberships_sql_statement += memberships_row
    return memberships_sql_statement

def correct_sql_statement(sql_statement):
##  The method gets unvalid SQL statement (as string) with extra ',' and '\n' also missing ';'
##  Returns a correct and valid SQL statement (as string).

    sql_statement = sql_statement[0:-2]
    sql_statement += ';'
    return sql_statement

if __name__== "__main__":
    main(sys.argv[1])