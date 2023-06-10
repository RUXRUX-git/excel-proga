import hashlib

import pandas as pd

import config

BUF_SIZE = 65536

def from_file(path):
    sha256 = hashlib.sha256()

    with open(path, 'rb') as f:
        while True:
            data = f.read(BUF_SIZE)
            if not data:
                break
            sha256.update(data)

    return sha256

def from_excel_file(path):
    sha256 = hashlib.sha256()
    df = pd.read_excel(path)
    data = str(df).encode()
    sha256.update(data)

    sha256.update(config.APP_SOLD)

    return sha256
