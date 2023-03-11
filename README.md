# sql2csv

## Features
* connects to a database
* reads all .sql files of current directory
* excecutes one after the other
* writes results to text (.csv) and Excel (.xslx) files

## Supported Databases
* PostgreSQL
* Oracle
* MS SQL 
* SQLite3

## Security Warning
ONLY USE READ-ONLY DB-USER ACCOUNTS via:<br/>
GRANT SELECT ON ALL TABLES IN SCHEMA schema_name TO username

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