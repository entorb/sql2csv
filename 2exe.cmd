@echo off
pyinstaller --onefile --console sql2csv.py
move dist\*.exe .\
