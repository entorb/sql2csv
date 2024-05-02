"""
Gen Checksum.

reads all .sql files of current directory
generate a sha256 hash for each file after adding a secret salt to it
writes hash to *.hash files
"""

import hashlib
from pathlib import Path

import sql2csv_credentials  # my credential file

# import sys
# pyinstaller --onefile --console

hash_salt = sql2csv_credentials.hash_salt


def gen_checksum(s: str, my_secret: str) -> str:
    """
    Calculate a sha256 checksum/hash of a string.

    add a "secret/salt" to the string to prevent others from being able to reproduce the
    checksum without knowing the secret
    """
    m = hashlib.sha256()
    m.update((s + my_secret).encode("utf-8"))
    return m.hexdigest()


if __name__ == "__main__":
    for filepath in Path().glob("*.sql"):
        print(f"File: {filepath}")
        # not set newline type here, it might be \n or \r\n
        with filepath.open(encoding="utf-8") as fh:
            cont = fh.read()

        checksum = gen_checksum(s=cont, my_secret=hash_salt)

        (filename, file_ext) = (filepath.stem, filepath.suffix)
        file_out = Path(filename + ".hash")
        with file_out.open(mode="w", encoding="utf-8", newline="\n") as fh:
            fh.write(checksum)
