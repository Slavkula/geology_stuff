import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox

def select_excel_file():  #от
    file_path = filedialog.askopenfilename()
    if file_path:
        excel_file_label.config(text=file_path)

def select_output_folder():
    folder_path = filedialog.askdirectory()
    if folder_path:
        output_folder_label.config(text=folder_path)

def convert_to_csv():
    global csv_files_list
    file_path = excel_file_label.cget("text")
    save_directory = output_folder_label.cget("text")

    try:
        df = pd.read_excel(file_path, sheet_name='Лист3')  # Изменяй на имя эксель листа
        for i in range(df.shape[1]):
            new_df = df.iloc[:, [0, i]]
            file_name = df.columns[i]
            csv_file_path = os.path.join(save_directory, f"{file_name}.csv")
            new_df.to_csv(csv_file_path, sep=';', index=False, encoding='utf-8-sig')
            csv_files_list.append(csv_file_path)  # Добавляем путь к каждому созданному файлу CSV
    except Exception as e:
        messagebox.showerror("Ошибка", str(e))

def delete_first_line():
    for file_path in csv_files_list:
        with open(file_path, 'r', encoding='utf-8') as file:
            lines = file.readlines()
        with open(file_path, 'w', encoding='utf-8') as file:
            file.writelines(lines[1:])  # Записываем все строки, кроме первой

root = tk.Tk()
root.title("Конвертер Excel в CSV")

global csv_files_list  # Глобальный список для хранения путей к файлам CSV
csv_files_list = []

excel_file_label = tk.Label(root, text="Выберите файл Excel:")
excel_file_label.pack()
excel_file_button = tk.Button(root, text="Выбрать файл", command=select_excel_file)
excel_file_button.pack()

output_folder_label = tk.Label(root, text="Выберите папку для сохранения файлов CSV:")
output_folder_label.pack()
output_folder_button = tk.Button(root, text="Выбрать папку", command=select_output_folder)
output_folder_button.pack()

convert_button = tk.Button(root, text="Преобразовать в CSV", command=convert_to_csv)
convert_button.pack()

delete_first_line_button = tk.Button(root, text="Удалить первую строку в файлах CSV", command=delete_first_line)
delete_first_line_button.pack()

text = """
Программа копирует столбцы с температурами из экселя. Берутся столбцы:
AA, AB, AC так далее до конца файла.
1.  Выбираешь эксель, из которого надо взять данные.
2.  Выбираешь папку, в которую нужно сохранить твои .csv файлы.
3.  Удаляешь кнопкой первые строки. Если файлов много, то может занять какое-то время.
"""
text_label = tk.Label(root, text=text)
text_label.pack() #до

root.mainloop()
