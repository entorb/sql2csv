#!/usr/bin/python
"""
source at
https://github.com/entorb/sql2csv

## Features
* connect to a database
* read all .sql files of current directory
* execute one after the other
* export results set as text (.csv) and Excel (.xslx)

## Supported Databases
* PostgreSQL
* Oracle
* MS SQL 
* SQLite3

## TODOs
- [x] scan SQL for dangerous commands like DROP/DELETE (incomplete!)
- [x] Limits the max number of returned rows via limit on cells = columns * rows
- [x] hashing of SQL files to prevent modification
- [ ] Excel: autosize column width
- [ ] use [sqlparse](https://sqlparse.readthedocs.io/en/latest/api/) to remove comments from SQL

## **SECURITY WARNING:** Only use read-only db-user accounts!
example for PostgreSQL<br/>
`GRANT SELECT ON ALL TABLES IN SCHEMA schema_name TO username`<br/>
`GRANT USAGE ON SCHEMA schema_name TO username`

## Requirements
### Oracle
Oracle Instant Client - Basic Light Package
from
https://www.oracle.com/database/technologies/instant-client/winx64-64-downloads.html
download, unzip and add dir to path

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

import csv  # for .csv writing
import openpyxl  # for Excel writing
import hashlib  # for sha256 checksums
import decimal

# DB drivers
import sqlite3
import pyodbc  # ODBC driver for MS SQL and others  - pip install pyodbc
import cx_Oracle  # driver for Oracle - pip install cx_Oracle
import psycopg2  # PostgreSQL - pip install psycopg2

# read my credential file
from sql2csv_credentials import credentials, hash_salt

# IDEA: move settings to .ini
cnt_max_cells = 100000
csv_format_datetime = "%Y-%m-%d_%H:%M:%S"
csv_format_date = "%Y-%m-%d"
csv_delimiter = "\t"  # \t ; ,
csv_quotechar = '"'
csv_newline = "\n"


# helper functions
def remove_old_output_files(fileBaseName: str):
    """
    remove output files prior to re-creation
    this is done prior to hash check to ensure that there is no output file in case the hash is bad
    """
    for ext in (".csv", ".xlsx"):
        outfile = fileBaseName + ext
        if os.path.isfile(outfile):
            os.remove(outfile)


# database access functions
def connect():
    """ connect to the database server """
    connection = None
    cursor = None
    # try:
    if credentials['db_type'] == 'postgres':
        connection = psycopg2.connect(host=credentials['host'],
                                      port=credentials['port'],
                                      database=credentials['database'],
                                      user=credentials['user'],
                                      password=credentials['password'])

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


def execute_sql(sql: str) -> list:
    """ excecute SQL statement and return results as list """
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
                    f"WARNING: too many values, stopped after {cnt_max_cells} values")
    except (Exception) as error:
        cursor.execute('ROLLBACK')
        print(error)
    return results


def sql_check_danger(sql: str):
    """
    raises Exception in case SQL contains dangerous commands
    Warning: incomplete, so not rely on this, use read only DB user too!
    """
    # bad = False
    bad_words = (
        'create',
        'alter',
        'drop',
        'grant',
        'revoke',
        'insert',
        'update',
        'delete',
        'truncate',
        'commit'
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


# export functions
def sql2csv(results: list, outfilename: str):
    """ write results into csv file """
    outfile = outfilename + '.csv'
    with open(outfile, mode='w', encoding='utf-8', newline=csv_newline) as fh:
        csvwriter = csv.writer(
            fh, delimiter=csv_delimiter, quotechar=csv_quotechar)
        colnames = results[0]  # header row
        csvwriter.writerow(colnames)
        for k in range(1, len(results)):
            row = results[k]
            row_str = []
            for value in row:
                value_str = convert_value_to_string(
                    value, remove_linebreaks=True, trim=True)
                row_str.append(value_str)
            csvwriter.writerow(row_str)


def sql2xlsx(results: list, outfilename: str):
    """ write results into Excel .xlsx file """
    outfile = outfilename + '.xlsx'
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
    """ convert SQL field types to strings, used in sql2csv """
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
        value_str = value.strftime(csv_format_datetime)
    elif t == datetime.date:
        value_str = value.strftime(csv_format_date)
    elif t == bool:
        if value == True:
            value_str = 1
        if value == False:
            value_str = 0
    elif t == decimal.Decimal:
        value_str = str(value)
    elif t == datetime.timedelta:
        value_str = str(value)
    else:
        print(f"unhandled column type: {t}")
        quit()
    return value_str


# checksum functions
def gen_checksum(s: str, my_secret: str) -> str:
    """
    calculate a sha256 checksum/hash of a string
    add a "secret/salt" to the string to prevent others from being able to reproduce the checksum without knowing the secret
    """
    m = hashlib.sha256()
    m.update((s + my_secret).encode('utf-8'))
    return m.hexdigest()


def check_for_valid_hashfile(sql: str, fileBaseName: str) -> bool:
    """
    return False if .hash file is missing or contains a wrong hash
    """
    valid = True
    filename = fileBaseName + ".hash"
    if valid:
        if not os.path.exists(filename):
            valid = False
            print(f"ERROR: missing checksum file '{filename}'")

    if valid:
        with open(fileBaseName+".hash", mode='r', encoding='utf-8', newline='\n') as fh:
            checksum_file = fh.read()
        checksum_calc = gen_checksum(s=sql, my_secret=hash_salt)
        if checksum_file != checksum_calc:
            valid = False
            print(f"ERROR: checksum missmatch")

    return valid


if __name__ == '__main__':
    (connection, cursor) = connect()
    for filename in glob.glob("*.sql"):
        print(f'File: {filename}')
        (fileBaseName, fileExtension) = os.path.splitext(filename)

        remove_old_output_files(fileBaseName)

        with open(filename, mode='r', encoding='utf-8') as fh:
            sql = fh.read()

        if hash_salt != "":  # perform hash check
            ret = check_for_valid_hashfile(
                sql=sql, fileBaseName=fileBaseName)
            if ret != True:
                continue

        sql_check_danger(sql=sql)
        results = execute_sql(sql=sql)
        # if len(results) > 1:
        sql2csv(results=results, outfilename=fileBaseName)
        sql2xlsx(results=results, outfilename=fileBaseName)

    if (connection):
        cursor.close()
        connection.close()
        print("Database connection closed")
