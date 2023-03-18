# sql2csv

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

* [x] scan SQL for dangerous commands like DROP/DELETE (incomplete!)
* [x] Limits the max number of returned rows via limit on cells = columns * rows
* [x] hashing of SQL files to prevent modification
* [ ] Excel: autosize column width
* [ ] use [sqlparse](https://sqlparse.readthedocs.io/en/latest/api/) to remove comments from SQL

## **SECURITY WARNING:** Only use read-only db-user accounts

example for PostgreSQL

```sql
GRANT SELECT ON ALL TABLES IN SCHEMA schema_name TO username
GRANT USAGE ON SCHEMA schema_name TO username
```

## Requirements

### Oracle

Oracle Instant Client - Basic Light Package
from
<https://www.oracle.com/database/technologies/instant-client/winx64-64-downloads.html>
download, unzip and add dir to path

### MS SQL

Microsoft ODBC Driver for SQL Server
from
<https://docs.microsoft.com/en-us/sql/connect/odbc/download-odbc-driver-for-sql-server?view=sql-server-ver15>
install
