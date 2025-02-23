import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, simpledialog
import time

def select_excel_file():
    root = tk.Tk()
    root.withdraw()  
    file_path = filedialog.askopenfilename(title="Выберите файл Excel", filetypes=[("Excel files", "*.xlsx;*.xls")])
    return file_path

def select_sheet_number(file_path):
    df = pd.ExcelFile(file_path)
    sheet_count = len(df.sheet_names)
    sheet_number = simpledialog.askinteger("Выбор листа", f"Введите номер листа (1-{sheet_count}):", minvalue=1, maxvalue=sheet_count)
    return sheet_number

def convert_to_csv(file_path, sheet_number):
    csv_files_list = []
    folder_path = os.path.dirname(file_path)
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_number-1)  
        for i in range(df.shape[1]):
            new_df = df.iloc[:, [0, i]]
            file_name = df.columns[i]
            csv_file_path = os.path.join(folder_path, f"{file_name}.csv")
            new_df.to_csv(csv_file_path, sep=';', index=False, encoding='utf-8-sig')
            csv_files_list.append(csv_file_path)
    except Exception as e:
        print("Ошибка:", str(e))
    return csv_files_list

def delete_first_line(csv_files_list):
    for csv_file_path in csv_files_list:
        with open(csv_file_path, 'r', encoding='utf-8') as file:
            lines = file.readlines()
        with open(csv_file_path, 'w', encoding='utf-8') as file:
            file.writelines(lines[1:])

def remove_empty_semicolon_rows(csv_files_list):
    for csv_file_path in csv_files_list:
        with open(csv_file_path, 'r', encoding='utf-8') as file:
            lines = file.readlines()
        
        with open(csv_file_path, 'w', encoding='utf-8') as file:
            for line in lines:
                if ';' in line and line.strip().endswith(';'):
                    continue
                file.write(line)

file_path = select_excel_file()
if file_path:
    sheet_number = select_sheet_number(file_path)
    if sheet_number:
        csv_files_list = convert_to_csv(file_path, sheet_number)
        delete_first_line(csv_files_list)
        remove_empty_semicolon_rows(csv_files_list)
        print("Все файлы успешно преобразованы и обработаны.")
    else:
        print("Лист не был выбран.")
else:
    print("Файл не был выбран.")

time.sleep(30)
