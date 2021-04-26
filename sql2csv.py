#!/usr/bin/python
"""
source at
https://github.com/entorb/sql2csv

* connects to a database
* reads all .sql files of current directory
* excecutes one after the other
* writes results to text (.csv) and Excel (.xslx) files

## Supported Databases
* PostgreSQL
* Oracle
* MS SQL 

## Security Warning
ONLY USE READ-ONLY DB-USER ACCOUNTS via:
GRANT SELECT ON ALL TABLES IN SCHEMA schema_name TO username

## Requirements
### Oracle
Oracle Instant Client - Basic Light Package
from
https://www.oracle.com/database/technologies/instant-client/winx64-64-downloads.html
download and unzip and add dir to path

### MS SQL
Microsoft ODBC Driver for SQL Server
from
https://docs.microsoft.com/en-us/sql/connect/odbc/download-odbc-driver-for-sql-server?view=sql-server-ver15
install
"""

# Convert to .exe via
# pyinstaller --onefile --console sql2csv.py

import os
import re
import glob
import datetime

import openpyxl  # Excel

import pyodbc  # ODBC
import cx_Oracle  # alternative library for oracle
import psycopg2  # PostgreSQL
import sqlite3


import my_credentials  # my credential file
# for sqlite3 only key 'database' -> path to file database is required
# example
# credentials_1 = {'db_type' : 'mssql', 'host': 'myHost', 'port': 5432,
#                  'database': 'myDB', 'user': 'myUser', 'password': 'myPwd'}
credentials = my_credentials.credentials_1

# supported db_types:
# - oracle
# - postgres
# - mssql
# - sqlite3

cnt_max_cells = 100000


def connect():
    """ Connect to the database server """
    connection = None
    cursor = None
    # try:
    if credentials['db_type'] == 'postgres':
        connection = psycopg2.connect(**credentials)
    elif credentials['db_type'] == 'sqlite3':
        # here database is the path to the database file
        connection = sqlite3.connect(credentials['database'])
        credentials['host'] = 'localhost'
    elif credentials['db_type'] == 'oracle':
        # connection = pyodbc.connect(
        #     f"DRIVER={{ORACLE ODBC DRIVER}};SERVER=tcp:{credentials['host']},{credentials['port']};DATABASE={credentials['database']};UID={credentials['user']};PWD={credentials['password']}"
        # )

        connection = cx_Oracle.connect(
            credentials['user'],
            credentials['password'],
            f"{credentials['host']}:{credentials['port']}/{credentials['database']}",
            encoding="UTF-8"
        )
    elif credentials['db_type'] == 'mssql':
        connection = pyodbc.connect(
            f"DRIVER={{ODBC Driver 17 for SQL Server}};SERVER=tcp:{credentials['host']},{credentials['port']};DATABASE={credentials['database']};UID={credentials['user']};PWD={credentials['password']}"
        )
    else:
        raise Exception(
            f"ERROR: unsupported db_type: '{credentials['db_type']}'")

    cursor = connection.cursor()
    print(
        f"connected to database {credentials['database']} on host {credentials['host']}")
    # except (Exception) as error:
    #     print("Error while connecting to Database", error)
    return connection, cursor


def sql_check_danger(sql: str):
    """ raises Exception in case SQL contains dangerous words """
    # bad = False
    bad_words = (
        'create',
        'alter',
        'drop',
        'insert',
        'update',
        'delete',
        'truncate'
    )
    res = re.match(r'\b(' + '|'.join(bad_words) + r')\b', sql.lower())
    if res != None:
        # bad = True
        raise Exception(f"ERROR: dangerous SQL: \n{sql}")
        # sys.exit(1)
    res = re.match(r'[^\b](select)\b', sql.lower())
    if res != None:
        # bad = True
        raise Exception(f"ERROR: not valid SQL: \n{sql}")
    # return bad


def execute_sql(sql: str) -> list:
    """ Excecute SQL statement and return results as list """
    results = []
    try:
        cursor.execute(sql)
        colnames = [desc[0] for desc in cursor.description]
        cnt_columns = len(colnames)
        results.append(colnames)
        cnt = 0
        for row in cursor:
            results.append(row)
            cnt += cnt_columns
            if cnt > cnt_max_cells:
                raise Exception(
                    f"ERROR: too many values, stopped after {cnt_max_cells} values")
    except (Exception) as error:
        cursor.execute('ROLLBACK')
        print(error)
    return results


def sql2csv(results: list, outfilename: str):
    """ Write results into csv file """
    outfile = outfilename + '.csv'
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


def sql2xlsx(results: list, outfilename: str):
    """ Write results into xlsx file """
    outfile = outfilename + '.xlsx'
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
    """ converts SQL field types to strings, used in sql2csv """
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
        sql_check_danger(sql=sql)
        results = execute_sql(sql=sql)
        if len(results) > 1:
            sql2csv(results=results, outfilename=fileBaseName)
            sql2xlsx(results=results, outfilename=fileBaseName)

    if (connection):
        cursor.close()
        connection.close()
        print("Database connection closed")
