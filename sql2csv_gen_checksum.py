#!/usr/bin/python
import glob
import os
# import sys
import hashlib

# pyinstaller --onefile --console

"""
reads all .sql files of current directory
generate a sha256 hash for each file after adding a secret salt to it
writes hash to *.hash files
"""

import sql2csv_credentials  # my credential file
hash_salt = sql2csv_credentials.hash_salt


def gen_checksum(s: str, my_secret: str) -> str:
    """ 
    calculates a sha256 checksum/hash of a string 
    add a "secret/salt" to the string to prevent others from being able to reproduce the checksum without knowing the secret
    """
    m = hashlib.sha256()
    m.update((s + my_secret).encode('utf-8'))
    return m.hexdigest()


if __name__ == '__main__':
    for filename in glob.glob("*.sql"):
        print(f'File: {filename}')
        fh = open(filename, mode='r', encoding='utf-8', newline='\n')
        cont = fh.read()
        fh.close()

        checksum = gen_checksum(s=cont, my_secret=hash_salt)

        (fileBaseName, fileExtension) = os.path.splitext(filename)
        fileOut = fileBaseName+'.hash'
        fh = open(fileOut, mode='w', encoding='utf-8', newline='\n')
        fh.write(checksum)
        fh.close()
