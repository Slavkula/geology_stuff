import win32com.client as win32
import tkinter as tk
from tkinter import messagebox


def process_data():
    try:
        col_I = entry_col_I.get().upper()  
        col_BK = entry_col_BK.get().upper()
        start_row = int(entry_start_row.get())
        end_row = int(entry_end_row.get())

        excel = win32.GetObject(Class="Excel.Application")
        workbook = excel.ActiveWorkbook
        sheet = workbook.ActiveSheet

        for row in range(start_row, end_row + 1):
            # Получаем значение из столбца I
            cell_I = sheet.Cells(row, sheet.Columns(col_I).Column)
            value_I = cell_I.Value

            if not isinstance(value_I, (int, float)):
                continue

            if value_I != 100:
                difference = 100 - value_I

                cell_BK = sheet.Cells(row, sheet.Columns(col_BK).Column)
                value_BK = cell_BK.Value if isinstance(cell_BK.Value, (int, float)) else 0

                new_value_BK = value_BK + difference
                cell_BK.Value = new_value_BK

    except Exception as e:
        messagebox.showerror("Ошибка", f"Произошла ошибка: {e}")

root = tk.Tk()
root.title("Убрать красные")

tk.Label(root, text="Столбец с красными:").grid(row=0, column=0, padx=10, pady=10)
entry_col_I = tk.Entry(root)
entry_col_I.grid(row=0, column=1, padx=10, pady=10)
entry_col_I.insert(0, "")  # Значение по умолчанию

tk.Label(root, text="Столбец для замены:").grid(row=1, column=0, padx=10, pady=10)
entry_col_BK = tk.Entry(root)
entry_col_BK.grid(row=1, column=1, padx=10, pady=10)
entry_col_BK.insert(0, "")  # Значение по умолчанию

tk.Label(root, text="Начать на строке:").grid(row=2, column=0, padx=10, pady=10)
entry_start_row = tk.Entry(root)
entry_start_row.grid(row=2, column=1, padx=10, pady=10)
entry_start_row.insert(0, "")  # Значение по умолчанию

tk.Label(root, text="Закончить на строке:").grid(row=3, column=0, padx=10, pady=10)
entry_end_row = tk.Entry(root)
entry_end_row.grid(row=3, column=1, padx=10, pady=10)
entry_end_row.insert(0, "")  # Значение по умолчанию

btn_process = tk.Button(root, text="Обработать", command=process_data)
btn_process.grid(row=4, column=0, columnspan=2, pady=10)

root.mainloop()
