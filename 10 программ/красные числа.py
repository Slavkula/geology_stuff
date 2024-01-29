import win32com.client
import tkinter as tk
from tkinter import filedialog, simpledialog

def run_macro():
    file_path = file_path_var.get()
    start_value = start_var.get()
    end_value = end_var.get()
    bf_value = bf_var.get()
    a_value = a_var.get()
    if file_path and start_value and end_value and bf_value and a_value:
        xl = win32com.client.Dispatch("Excel.Application")
        xl.Visible = True  # Раскомментируйте эту строку, если вы хотите видеть процесс выполнения макроса

        wb = xl.Workbooks.Open(file_path)
        ws = wb.ActiveSheet

        # Создаем макрос в VBA с использованием введенных значений
        macro_code = f"""
        Sub CustomLoop()
            Dim ws As Worksheet
            Dim i As Long
            Set ws = ActiveSheet

            For i = {start_value} To {end_value}
                Dim originalValue As Variant
                Dim adjustedValue As Double
                Dim bfValue As Double

                originalValue = ws.Cells(i, "{a_value}").Value
                If IsNumeric(originalValue) Then
                    adjustedValue = 100 - originalValue
                    bfValue = ws.Cells(i, "{bf_value}").Value
                    ws.Cells(i, "{bf_value}").Value = bfValue + adjustedValue
                End If
            Next i
        End Sub
        """

        xlmodule = wb.VBProject.VBComponents.Add(1)
        xlmodule.CodeModule.AddFromString(macro_code)

        # Запускаем созданный макрос
        xl.Run("CustomLoop")

        # Сохраняем изменения и закрываем файл
        wb.Save()
        #xl.Quit()

        result_label.config(text="Макрос успешно выполнен")
    else:
        result_label.config(text="Путь к файлу или значения для цикла не указаны")

def select_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if file_path:
        file_path_var.set(file_path)
        result_label.config(text="")
        select_button.config(bg="green")  


root = tk.Tk()
root.title("Убрать красные числа")
root.option_add('*Font', 'Helvetica 20')  

file_path_var = tk.StringVar()
start_var = tk.StringVar()
end_var = tk.StringVar()
bf_var = tk.StringVar()
a_var = tk.StringVar()

file_label = tk.Label(root, text="Выберите файл Excel:", font=('Helvetica', 20))  
file_label.pack(pady=20)

select_button = tk.Button(root, text="Выбрать файл", command=select_file, font=('Helvetica', 20))  
select_button.pack()

start_frame = tk.Frame(root)  
start_frame.pack(pady=20)

start_label = tk.Label(start_frame, text="Начало:", font=('Helvetica', 20))  
start_label.pack(side="left")

start_entry = tk.Entry(start_frame, textvariable=start_var, width=5, font=('Helvetica', 20))  
start_entry.pack(side="right")

end_frame = tk.Frame(root)  
end_frame.pack(pady=20)

end_label = tk.Label(end_frame, text="Конец:", font=('Helvetica', 20))  
end_label.pack(side="left")

end_entry = tk.Entry(end_frame, textvariable=end_var, width=5, font=('Helvetica', 20))  
end_entry.pack(side="right")

bf_frame = tk.Frame(root)  
bf_frame.pack(pady=20)

bf_label = tk.Label(bf_frame, text="Делать замены в столбце:", font=('Helvetica', 20))  
bf_label.pack(side="left")

bf_entry = tk.Entry(bf_frame, textvariable=bf_var, width=5, font=('Helvetica', 20))  
bf_entry.pack(side="right")

a_frame = tk.Frame(root)  
a_frame.pack(pady=20)

a_label = tk.Label(a_frame, text="Красные числа в столбце:", font=('Helvetica', 20))  
a_label.pack(side="left")

a_entry = tk.Entry(a_frame, textvariable=a_var, width=5, font=('Helvetica', 20))  
a_entry.pack(side="right")

run_button = tk.Button(root, text="Запустить макрос", command=run_macro, font=('Helvetica', 20))  
run_button.pack(pady=20)

result_label = tk.Label(root, text="", font=('Helvetica', 20))  
result_label.pack()

root.mainloop()
