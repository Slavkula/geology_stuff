import openpyxl
import os

# Открываем файл 2.xlsx и считываем данные
wb2 = openpyxl.load_workbook('2.xlsx')
sheet2 = wb2.active

for i in range(70, 81):
    # Считываем данные
    data = [sheet2[f'Q{i}'].value, sheet2[f'R{i}'].value, sheet2[f'T{i}'].value, sheet2[f'U{i}'].value]

    # Получаем значения для имени файла
    filename_part1 = sheet2[f'H{i}'].value
    filename_part2 = sheet2[f'I{i}'].value

    # Создаем новый файл Excel и записываем данные
    wb_new = openpyxl.Workbook()
    sheet_new = wb_new.active
    sheet_new['B9'] = data[0]
    sheet_new['B10'] = data[1]
    sheet_new['B11'] = data[2]
    sheet_new['D11'] = data[3]

    # Сохраняем изменения в новом файле Excel
    wb_new.save(f'{filename_part1}-{filename_part2}.xlsx')
