import tkinter as tk
from tkinter import filedialog
import win32com.client

def main():
    root = tk.Tk()
    root.title("Создаём влажность")

    file_button = tk.Button(root, text="Обзор", command=lambda: select_file(), height=2, width=20, font=("Helvetica", 16))
    file_button.grid(row=0, column=1)

    tk.Label(root, text="Выбери эксель файл", font=("Helvetica", 16)).grid(row=0, column=0)

    list2_entry = tk.Entry(root, font=("Helvetica", 16), width=12)
    list2_entry.grid(row=1, column=1)

    tk.Label(root, text="Имя листа для вставки:", font=("Helvetica", 16)).grid(row=1, column=0)

    values_entry = [tk.Entry(root, font=("Helvetica", 16), width=12) for _ in range(4)]
    for i, entry in enumerate(values_entry):
        entry.grid(row=2+i, column=1)
        tk.Label(root, text=f"Наименование ИГЭ: {i+1}", font=("Helvetica", 16)).grid(row=2+i, column=0)

    t_entry = tk.Entry(root, font=("Helvetica", 16), width=12)
    t_entry.grid(row=6, column=1)

    tk.Label(root, text="Столбец для сверки ИГЭ:", font=("Helvetica", 16)).grid(row=6, column=0)

    y_entry = tk.Entry(root, font=("Helvetica", 16), width=12)
    y_entry.grid(row=7, column=1)

    tk.Label(root, text="Столбец для сверки влажности:", font=("Helvetica", 16)).grid(row=7, column=0)

    range_entry = tk.Entry(root, font=("Helvetica", 16), width=12)
    range_entry.grid(row=8, column=1)

    tk.Label(root, text="Диапазон в формате R29:O29:", font=("Helvetica", 16)).grid(row=8, column=0)

    lower_bound_entry = tk.Entry(root, font=("Helvetica", 16), width=12)
    lower_bound_entry.grid(row=9, column=1)

    tk.Label(root, text="Граница влажности от:", font=("Helvetica", 16)).grid(row=9, column=0)

    upper_bound_entry = tk.Entry(root, font=("Helvetica", 16), width=12)
    upper_bound_entry.grid(row=10, column=1)

    tk.Label(root, text="Граница влажности до:", font=("Helvetica", 16)).grid(row=10, column=0)

    run_button = tk.Button(root, text="Запустить", command=lambda: run_macro(), height=2, width=20, font=("Helvetica", 16))
    run_button.grid(row=11, column=1)

    file_path = None
    Excel = None
    Workbook = None
    Module = None

    def select_file():
        nonlocal file_path
        file_path = filedialog.askopenfilename()
        file_button.config(bg="green")

    def run_macro():
        nonlocal Excel, Workbook, Module
        Excel = win32com.client.Dispatch("Excel.Application")
        Excel.Visible = True
        Workbook = Excel.Workbooks.Open(file_path)
        Module = Workbook.VBProject.VBComponents.Add(1)

        list2 = list2_entry.get()
        values = [entry.get() for entry in values_entry if entry.get()]
        t = t_entry.get()
        y = y_entry.get()
        range_ = range_entry.get()
        lower_bound = lower_bound_entry.get()
        upper_bound = upper_bound_entry.get()

        conditions = " Or ".join([f'Sheets("{list2}").Range("M" & i).Value = "{value}"' for value in values])

        macro = f"""
        Sub Dolbilka()
            Dim i As Integer
            For i = 1 To 10000
                Sheets("Лист1").Select
                If {conditions} Then
                    Do
                        Range("{range_}").Copy
                        Sheets("{list2}").Range("{t}" & i).PasteSpecial xlPasteValues
                    Loop While Sheets("{list2}").Range("{y}" & i).Value < {lower_bound} Or Sheets("{list2}").Range("{y}" & i).Value > {upper_bound}
                End If
            Next i
        End Sub
        """

        Module.CodeModule.AddFromString(macro)
        Excel.Application.Run("Dolbilka")
        Workbook.Save()

    root.mainloop()

if __name__ == "__main__":
    main()
