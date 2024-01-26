import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox

def select_excel_file():
    global file_path
    file_path = filedialog.askopenfilename()
    if file_path:
        excel_file_button.config(bg='green')

def select_output_folder():
    global folder_path
    folder_path = filedialog.askdirectory()
    if folder_path:
        output_folder_button.config(bg='green')

def convert_to_csv():
    global csv_files_list
    try:
        df = pd.read_excel(file_path)  # Берем данные из активного листа
        for i in range(df.shape[1]):
            new_df = df.iloc[:, [0, i]]
            file_name = df.columns[i]
            csv_file_path = os.path.join(folder_path, f"{file_name}.csv")
            new_df.to_csv(csv_file_path, sep=';', index=False, encoding='utf-8-sig')
            csv_files_list.append(csv_file_path)
    except Exception as e:
        messagebox.showerror("Ошибка", str(e))

def delete_first_line():
    for csv_file_path in csv_files_list:
        with open(csv_file_path, 'r', encoding='utf-8') as file:
            lines = file.readlines()
        with open(csv_file_path, 'w', encoding='utf-8') as file:
            file.writelines(lines[1:])

root = tk.Tk()
root.title("Конвертер Excel в CSV")

global csv_files_list
csv_files_list = []

file_path = ""
folder_path = ""

excel_file_button = tk.Button(root, text="Выбрать файл", command=select_excel_file, font=("Arial", 20), height=2, width=30)
excel_file_button.pack()

output_folder_button = tk.Button(root, text="Выбрать папку", command=select_output_folder, font=("Arial", 20), height=2, width=30)
output_folder_button.pack()

convert_button = tk.Button(root, text="Преобразовать в CSV", command=convert_to_csv, font=("Arial", 20), height=2, width=30)
convert_button.pack()

delete_first_line_button = tk.Button(root, text="Удалить первую строку в файлах CSV", command=delete_first_line, font=("Arial", 20), height=2, width=30)
delete_first_line_button.pack()

text = """
Программа копирует столбцы с температурами из экселя. Берутся столбцы:
AA, AB, AC так далее до конца файла.
1.  Выбираешь эксель, из которого надо взять данные.
2.  Выбираешь папку, в которую нужно сохранить твои .csv файлы.
3.  Удаляешь кнопкой первые строки. Если файлов много, то может занять какое-то время.
"""
text_label = tk.Label(root, text=text, font=("Arial", 20))
text_label.pack()

root.mainloop()
