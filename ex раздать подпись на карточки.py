import tkinter as tk
from win32com.client import Dispatch

def run_macro():
    cell = entry.get()
    xl = Dispatch("Excel.Application")
    wb = xl.ActiveWorkbook
    ws = wb.Worksheets(1)
    macro = f"""
    Sub CopyPictureToAllSheets(cell As String)
        Dim ws As Worksheet
        Dim pic As Shape
        Set pic = Selection.ShapeRange.Item(1)
        For Each ws In ActiveWorkbook.Worksheets
            ws.Activate
            pic.CopyPicture Appearance:=xlScreen, Format:=xlPicture
            Range(cell).Select
            ActiveSheet.Paste
        Next ws
    End Sub
    """
    module = ws.Application.VBE.ActiveVBProject.VBComponents.Add(1)
    module.CodeModule.AddFromString(macro)
    ws.Application.Run("CopyPictureToAllSheets", cell)
    ws.Application.VBE.ActiveVBProject.VBComponents.Remove(module)

root = tk.Tk()

label = tk.Label(root, text="Введите ячейку для вставки:", font=("Arial", 18))  # увеличиваем размер шрифта
entry = tk.Entry(root, font=("Arial", 18))  
button = tk.Button(root, text="Вставить и выполнить макрос", command=run_macro, font=("Arial", 18))  # увеличиваем размер шрифта
label.pack(pady=10)  
entry.pack(pady=10)  
button.pack(pady=10)  
root.mainloop()
