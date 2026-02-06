import time

import pandas as pd

start = time.time()
# Читаем файл
df = pd.read_excel('work_files/grafic.xlsx')

# Проверяем, что есть достаточно колонок
if len(df.columns) > 21:
    target_column = df.columns[22]

    # Ищем первое вхождение '0209'
    for idx, value in enumerate(df[target_column].astype(str)):
        if value == '0209':
            row = df.iloc[idx]

            print("Месторождение:", row.iloc[2] if len(row) > 2 else "")
            print("Цех:", row.iloc[3] if len(row) > 3 else "")
            print("Инв. номер:", row.iloc[4] if len(row) > 4 else "")
            print("Трубопровод:", row.iloc[5] if len(row) > 5 else "")
            break
    else:
        print("Строка с '0209' не найдена")
else:
    print("Недостаточно колонок в файле")

end = time.time()

print("Time = ", end - start)