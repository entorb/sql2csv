#!/usr/bin/python
"""
connects to a database
reads all .sql files of current directory
excecutes one after the other
results are written to .csv files

ONLY USE READ-ONLY DB-USER ACCOUNTS via:
GRANT SELECT ON SCHEMA :: [dbo] TO username
"""
# Convert to .exe via
# pyinstaller --onefile --console sql2csv_mssql.py


import os
import glob
import csv
import pyodbc
# pip install pyodbc==4.0.27
# to fix
# "ImportError: DLL load failed"
# use
# pip install pyodbc==4.0.27
# as proposed in https://github.com/mkleehammer/pyodbc/issues/663

credentials_MyDB1 = {'host': 'myHost', 'port': 1433,
                     'database': 'myDB', 'user': 'myUser', 'password': 'myPwd'}
credentials = credentials_MyDB1


def connect():
    """ Connect to the MSSQL database server """

    connection = pyodbc.connect(
        f"DRIVER={{ODBC Driver 17 for SQL Server}};SERVER=tcp:{credentials['host']},{credentials['port']};DATABASE={credentials['database']};UID={credentials['user']};PWD={credentials['password']}")
    cursor = connection.cursor()
    print(
        f"connected to database {credentials['database']} on host {credentials['host']}")
    return connection, cursor


def sql2csv(cursor, sql: str, outfile: str = 'out.csv'):
    """ Excecute SQL statement and write results into csv file """
    cursor.execute(sql)
    with open(outfile, 'w', newline='') as csvfile:
        writer = csv.writer(csvfile, dialect='excel',
                            delimiter=';', quoting=csv.QUOTE_NONNUMERIC)
        # column headers
        writer.writerow([x[0] for x in cursor.description])
        for row in cursor:
            writer.writerow(row)


if __name__ == '__main__':
    (connection, cursor) = connect()
    for filename in glob.glob("*.sql"):
        print(f'File: {filename}')
        fh = open(filename, "r")
        sql = fh.read()
        fh.close()
        (fileBaseName, fileExtension) = os.path.splitext(filename)
        sql2csv(cursor, sql, fileBaseName+'.csv')

    if (connection):
        cursor.close()
        connection.close()
        print("MSSQL connection closed")
