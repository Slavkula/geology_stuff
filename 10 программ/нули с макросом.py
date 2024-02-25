import tkinter as tk
from tkinter import messagebox
import win32com.client

def generate_macro():
    v_start = v_start_entry.get()
    v_end = v_end_entry.get()
    k = k_entry.get()
    u = u_entry.get()

    macro = f"""
    Sub noliki()
        Dim rng As Range
        Dim cell As Range
        Dim subCell As Range

        Set rng = ActiveSheet.Range("{v_start}:{v_end}")

        For Each cell In rng
            If cell.Value <> "" Then
                For Each subCell In ActiveSheet.Range("{k}" & cell.Row & ":{u}" & cell.Row)
                    If subCell.Value = "" Then
                        subCell.Value = 0
                        subCell.Font.Color = RGB(255, 255, 255)
                    End If
                Next subCell
            End If
        Next cell
    End Sub
    """

    Excel = win32com.client.GetActiveObject("Excel.Application")
    wb = Excel.ActiveWorkbook

    mod = wb.VBProject.VBComponents.Add(1)
    mod.CodeModule.AddFromString(macro)

    Excel.Run("noliki")

    wb.VBProject.VBComponents.Remove(mod)

root = tk.Tk()
root.option_add('*Font', 'TkDefaultFont 18')  

v_start_label = tk.Label(root, text="Проверка от:")
v_start_entry = tk.Entry(root, width=10)
v_end_label = tk.Label(root, text="Проверка до:")
v_end_entry = tk.Entry(root, width=10)
k_label = tk.Label(root, text="Вставка от:")
k_entry = tk.Entry(root, width=10)
u_label = tk.Label(root, text="Вставка до:")
u_entry = tk.Entry(root, width=10)
button = tk.Button(root, text="Выполнить", command=generate_macro)

v_start_label.grid(row=0, column=0)
v_start_entry.grid(row=0, column=1)
v_end_label.grid(row=1, column=0)
v_end_entry.grid(row=1, column=1)
k_label.grid(row=2, column=0)
k_entry.grid(row=2, column=1)
u_label.grid(row=3, column=0)
u_entry.grid(row=3, column=1)
button.grid(row=4, column=0, columnspan=2)

root.mainloop()
