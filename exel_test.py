import time

start = time.time()

import pandas as pd

df = pd.read_excel('work_files/grafic.xlsx', usecols='A:Z', nrows=100, dtype=str, engine='openpyxl')

csv = df.to_csv().split('\r\n', )
csv = [i.split(',') for i in csv]

row_index = 0

report_number = csv[row_index][23]

while report_number != '0209':
    row_index += 1

    report_number = csv[row_index][23]
    if row_index == 100:
        print("BREAK")
        break

print("Месторождение:", csv[row_index][3])
print("Цех:", csv[row_index][4])
print("Инв. номер:", csv[row_index][5])
print("Трубопровод:", csv[row_index][7])

end = time.time()

print("Time = ", end - start)