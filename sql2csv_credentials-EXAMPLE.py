"""
Example credential file.
"""

# set to hash_salt = "" if check of hash checksum is not required
# hash_salt = ""
hash_salt = "This salt is my secret!"

# supported db_types:
# - oracle
# - postgres
# - mssql
# - sqlite3

# supported db_type=sqlite3 only key 'database' -> path to file database is required

credentials = {
    "db_type": "mssql",
    "host": "myHost",
    "port": 5432,
    "database": "myDB",
    "user": "myUser",
    "password": "myPwd",
}
