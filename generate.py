import random
import string

import pandas as pd

import config

def make_students_excel(path):
    logins = []

    for _ in range(config.STUDENTS_COUNT):
        login = ''.join(random.choice(string.ascii_letters + string.digits) for _ in range(8))
        logins.append(login)

    df = pd.DataFrame({
        'ФИО': [''] * len(logins),
        'Логин': logins,
    })
    df.index = range(1, len(df) + 1)  # Чтобы нумерация была не с 0, а с 1

    df.to_excel(path)
