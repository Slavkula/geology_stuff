import win32com.client
import tkinter as tk
from tkinter import filedialog, Entry, Label, StringVar, Button

def select_file():
    file_path = filedialog.askopenfilename()
    file_path_var.set(file_path)
    if file_path:
        select_file_button.config(bg="green")
    else:
        select_file_button.config(bg="red")

def run_macro():
    file_path = file_path_var.get()
    g_value = g_var.get()
    s_value = s_var.get()
    m_value = m_var.get()
    values = [value.get() for value in value_vars]
    bg_value = bg_var.get()

    if file_path and g_value and s_value and m_value and all(values) and bg_value:
        xl = win32com.client.Dispatch("Excel.Application")
        xl.Visible = True

        wb = xl.Workbooks.Open(file_path)
        ws = wb.ActiveSheet

        macro_code = f"""
Sub Kikimora()
    Dim i As Integer
    Sheets("Лист1").Select
    Dim cellValue As String
    Range("{g_value}:{s_value}").Select
    Selection.Copy
    Sheets(1).Select
    For i = 1 To 20000
        cellValue = Sheets(1).Range("{m_value}" & i).Value ' Использование индекса первой вкладки
        If cellValue = "{values[0]}" Or cellValue = "{values[1]}" Or cellValue = "{values[2]}" Or cellValue = "{values[3]}" Then
            Range("{bg_value}" & i).Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
        End If
    Next i
End Sub
        """

        xlmodule = wb.VBProject.VBComponents.Add(1)
        xlmodule.CodeModule.AddFromString(macro_code)

        xl.Run("Kikimora")

        wb.Save()

root = tk.Tk()
root.title("Заполнение ведомости")
root.option_add('*Font', 'Helvetica 20')

file_path_var = StringVar()
g_var = StringVar()
s_var = StringVar()
m_var = StringVar()
value_vars = [StringVar() for _ in range(4)]
bg_var = StringVar()

select_file_button = Button(root, text="Выберите файл Excel", command=select_file, font=('Helvetica', 20))
select_file_button.pack()

input_frame = tk.Frame(root)
input_frame.pack()

file_label = Label(input_frame, text="Копировать с:", font=('Helvetica', 20))
file_label.grid(row=0, column=0, sticky="e")
g_entry = Entry(input_frame, textvariable=g_var, font=('Helvetica', 20), width=12)
g_entry.grid(row=0, column=1)
g_entry.insert(0, "G250")
g_entry.config(fg="grey")
g_entry.bind("<FocusIn>", lambda event, entry=g_entry: on_entry_click(event, entry))
g_entry.bind("<FocusOut>", lambda event, entry=g_entry: on_focus_out(event, entry))

file_label = Label(input_frame, text="Копировать до:", font=('Helvetica', 20))
file_label.grid(row=1, column=0, sticky="e")
s_entry = Entry(input_frame, textvariable=s_var, font=('Helvetica', 20), width=12)
s_entry.grid(row=1, column=1)
s_entry.insert(0, "S250")
s_entry.config(fg="grey")
s_entry.bind("<FocusIn>", lambda event, entry=s_entry: on_entry_click(event, entry))
s_entry.bind("<FocusOut>", lambda event, entry=s_entry: on_focus_out(event, entry))

def on_entry_click(event, entry):
    if entry.get() == entry.get() and entry.cget('fg') == "grey":
        entry.delete(0, "end")
        entry.config(fg="black")

def on_focus_out(event, entry):
    if entry.get() == "":
        entry.insert(0, entry.get())
        entry.config(fg="grey")

file_label = Label(input_frame, text="Сверка ИГЭ:", font=('Helvetica', 20))
file_label.grid(row=2, column=0, sticky="e")
m_entry = Entry(input_frame, textvariable=m_var, font=('Helvetica', 20), width=12)
m_entry.grid(row=2, column=1)

for i in range(4):
    label_text = f"ИГЭ {i+1}:"
    file_label = Label(input_frame, text=label_text, font=('Helvetica', 20))
    file_label.grid(row=i+3, column=0, sticky="e")
    value_entries = Entry(input_frame, textvariable=value_vars[i], font=('Helvetica', 20), width=12)
    value_entries.insert(0, "Введите ИГЭ")
    value_entries.config(fg="grey")
    value_entries.bind("<FocusIn>", lambda event, entry=value_entries: on_entry_click(event, entry))
    value_entries.bind("<FocusOut>", lambda event, entry=value_entries: on_focus_out(event, entry))
    value_entries.grid(row=i+3, column=1)

def on_entry_click(event, entry):
    if entry.get() == "Введите ИГЭ":
        entry.delete(0, "end")
        entry.config(fg="black")

def on_focus_out(event, entry):
    if entry.get() == "":
        entry.insert(0, "Введите ИГЭ")
        entry.config(fg="grey")

file_label = Label(input_frame, text="Вставка в:", font=('Helvetica', 20))
file_label.grid(row=7, column=0, sticky="e")
bg_entry = Entry(input_frame, textvariable=bg_var, font=('Helvetica', 20), width=12)
bg_entry.grid(row=7, column=1)

run_button = Button(root, text="Тык! Тык!", command=run_macro, font=('Helvetica', 20))
run_button.pack()

root.mainloop()
