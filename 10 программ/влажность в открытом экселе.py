import win32com.client
from tkinter import Tk, Label, Entry, Button, StringVar

def run_macro():
    Excel = win32com.client.Dispatch("Excel.Application")
    Excel.Visible = True 
    wb = Excel.ActiveWorkbook

    macro_vba_code = f"""
    Sub vlazhnost()
        Dim cellValue As String
        Sheets("Лист1").Select
        Range("{range_value.get()}").Select
        Selection.Copy
        Sheets(1).Select
        For i = {start_value.get()} To 10000
            cellValue = Range("{column_value.get()}" & i).Value
            If (cellValue = "{value1.get()}" Or cellValue = "{value2.get()}" Or cellValue = "{value3.get()}" Or cellValue = "{value4.get()}") And cellValue <> "" Then
                Range("{column_select.get()}" & i).Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
                Do While Range("{column_while.get()}" & i).Value < {lower_bound.get()} Or Range("{column_while.get()}" & i).Value > {upper_bound.get()}
                    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
                Loop
            End If
        Next i
    End Sub
    """

    ExcelModule = wb.VBProject.VBComponents.Add(1)
    ExcelModule.CodeModule.AddFromString(macro_vba_code)

    Excel.Application.Run('vlazhnost')
    wb.VBProject.VBComponents.Remove(ExcelModule)

root = Tk()
root.title("Создаём влажность")

range_value = StringVar()
column_value = StringVar()
value1 = StringVar()
value2 = StringVar()
value3 = StringVar()
value4 = StringVar()
column_while = StringVar()
lower_bound = StringVar()
upper_bound = StringVar()
start_value = StringVar()
column_select = StringVar()

Label(root, text="Данные копируем из", font=("Helvetica", 16)).grid(row=0)
Entry(root, textvariable=range_value, font=("Helvetica", 16), width=8).grid(row=0, column=1)

Label(root, text="ИГЭ проверяем по столбцу", font=("Helvetica", 16)).grid(row=1)
Entry(root, textvariable=column_value, font=("Helvetica", 16), width=4).grid(row=1, column=1)

Label(root, text="ИГЭ 1", font=("Helvetica", 16)).grid(row=2)
Entry(root, textvariable=value1, font=("Helvetica", 16), width=4).grid(row=2, column=1)

Label(root, text="ИГЭ 2", font=("Helvetica", 16)).grid(row=3)
Entry(root, textvariable=value2, font=("Helvetica", 16), width=4).grid(row=3, column=1)

Label(root, text="ИГЭ 3", font=("Helvetica", 16)).grid(row=4)
Entry(root, textvariable=value3, font=("Helvetica", 16), width=4).grid(row=4, column=1)

Label(root, text="ИГЭ 4", font=("Helvetica", 16)).grid(row=5)
Entry(root, textvariable=value4, font=("Helvetica", 16), width=4).grid(row=5, column=1)

Label(root, text="Влажность проверяeм по", font=("Helvetica", 16)).grid(row=6)
Entry(root, textvariable=column_while, font=("Helvetica", 16), width=4).grid(row=6, column=1)

Label(root, text="Влажность от", font=("Helvetica", 16)).grid(row=7)
Entry(root, textvariable=lower_bound, font=("Helvetica", 16), width=4).grid(row=7, column=1)

Label(root, text="Влажность до", font=("Helvetica", 16)).grid(row=8)
Entry(root, textvariable=upper_bound, font=("Helvetica", 16), width=4).grid(row=8, column=1)

Label(root, text="Вставка со строки", font=("Helvetica", 16)).grid(row=9)
Entry(root, textvariable=start_value, font=("Helvetica", 16), width=4).grid(row=9, column=1)

Label(root, text="Столбец для вставки", font=("Helvetica", 16)).grid(row=10)
Entry(root, textvariable=column_select, font=("Helvetica", 16), width=4).grid(row=10, column=1)

Button(root, text="Обработать", command=run_macro, font=("Helvetica", 16)).grid(row=11, column=1)

root.mainloop()
