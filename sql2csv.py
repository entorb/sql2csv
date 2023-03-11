#!/usr/bin/python
"""
connects to a database
reads all .sql files of current directory
excecutes one after the other
results are written to .csv files

ONLY USE READ-ONLY DB-USER ACCOUNTS via:
GRANT SELECT ON ALL TABLES IN SCHEMA schema_name TO username

Requirements
Oracle Instant Client - Basic Light Package
from
https://www.oracle.com/database/technologies/instant-client/winx64-64-downloads.html
download and unzip and add dir to path

"""

# Convert to .exe via
# pyinstaller --onefile --console sql2csv_oracle.py

# FIXME: this security check is not working
# if (sql1.find('insert') or sql1.find('update') or sql1.find('delete') or sql1.find('drop') or sql1.find('grant')):
#     print ("invalid SQL")
#     sys.exit(1)


import os
import glob
import cx_Oracle
import psycopg2
import datetime
import openpyxl
import my_credentials
credentials = {'host': 'myHost', 'port': 5432,
               'database': 'myDB', 'user': 'myUser', 'password': 'myPwd'}

# credentials = my_credentials.credentials
# credentials = my_credentials.credentials


# db_type = 'oracle'
db_type = 'postgres'


def connect():
    """ Connect to the database server """
    connection = None
    cursor = None
    # try:
    if db_type == 'oracle':
        connection = cx_Oracle.connect(
            credentials['user'],
            credentials['password'],
            f"{credentials['host']}:{credentials['port']}/{credentials['database']}",
            encoding="UTF-8"
        )
    if db_type == 'postgres':
        connection = psycopg2.connect(**credentials)

    cursor = connection.cursor()
    print(
        f"connected to database {credentials['database']} on host {credentials['host']}")
    # except (Exception) as error:
    #     print("Error while connecting to Database", error)
    return connection, cursor


def execute_sql(sql: str) -> list:
    """ Excecute SQL statement and return results as list """
    results = []
    try:
        cursor.execute(sql)
        colnames = [desc[0] for desc in cursor.description]
        results.append(colnames)
        cnt = 0
        for row in cursor:
            results.append(row)
            cnt += 1
            if cnt > 1000:
                raise Exception(
                    "ERROR: too many rows, stopped after 1000 rows")
    except (Exception) as error:
        cursor.execute('ROLLBACK')
        print(error)
    return results


def sql2csv(results: list, outfile: str = 'out.csv'):
    """ Write results into csv file """
    if os.path.isfile(outfile):
        os.remove(outfile)

    if len(results) <= 1:
        return
    with open(outfile, mode='w', encoding='utf-8', newline='\n') as fh:
        colnames = results[0]  # header row
        fh.write("\t".join(colnames))
        fh.write("\n")
        for k in range(1, len(results)):
            row = results[k]
            row_str = []
            for value in row:
                value_str = convert_value_to_string(
                    value, remove_linebreaks=True, trim=True)
                row_str.append(value_str)
            fh.write("\t".join(row_str))
            fh.write("\n")


def sql2xlsx(results: list, outfile: str = 'out.xlsx'):
    """ Write results into xlsx file """
    if os.path.isfile(outfile):
        os.remove(outfile)
    if len(results) <= 1:
        return
    workbookOut = openpyxl.Workbook()
    sheetOut = workbookOut.active

    colnames = results.pop(0)  # header row
    i = 1
    j = 1
    for value in colnames:
        # note: excel index start here with 1
        cellOut = sheetOut.cell(row=i, column=j)
        cellOut.value = value
        cellOut.font = openpyxl.styles.Font(bold=True)
        j += 1

    # contents
    i = 2
    for k in range(1, len(results)):
        row = results[k]
#    for row in results:
        j = 1
        for value in row:
            cellOut = sheetOut.cell(row=i, column=j)
            cellOut.value = value
            j += 1
        i += 1

    # TODO: autosize columns width
    # V1 from https://stackoverflow.com/questions/39529662/python-automatically-adjust-width-of-an-excel-files-columns
    # V2 # from https://izziswift.com/openpyxl-adjust-column-width-size/

    workbookOut.save(outfile)


def convert_value_to_string(value, remove_linebreaks: bool = True, trim: bool = True) -> str:
    value_str = ""
    t = type(value)
    if value == None:
        value_str = ""
    elif t == str:
        value_str = value
        if remove_linebreaks:
            value_str = value_str.replace("\r\n", " ").replace(
                "\n", " ").replace("\r", " ")
        if trim:
            value_str = value_str.strip()
            value_str = value_str.replace("\t", " ")  # tabs
            value_str = value_str.replace("  ", " ").replace(
                "  ", " ")  # multiple spaces

    elif t == int:
        value_str = str(value)
    elif t == datetime.datetime:
        value_str = value.strftime("%Y-%m-%d_%H:%M:%S")
    elif t == datetime.date:
        value_str = value.strftime("%Y-%m-%d")
    else:
        print(f"unhandled column type: {t}")
    return value_str


if __name__ == '__main__':
    (connection, cursor) = connect()
    for filename in glob.glob("*.sql"):
        print(f'File: {filename}')
        (fileBaseName, fileExtension) = os.path.splitext(filename)

        with open(filename, mode='r', encoding='utf-8') as fh:
            sql = fh.read()
        results = execute_sql(sql=sql)
        if len(results) > 1:
            sql2csv(results=results, outfile=fileBaseName+".csv")
            sql2xlsx(results=results, outfile=fileBaseName+".xlsx")

    if (connection):
        cursor.close()
        connection.close()
        print("Database connection closed")
