@echo off
pyinstaller --onefile --console sql2csv.py --hiddenimport=sql2csv_credentials
move dist\*.exe .\
