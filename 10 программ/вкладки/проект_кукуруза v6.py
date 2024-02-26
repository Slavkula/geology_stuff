import tkinter as tk
from tkinter import ttk
import pandas as pd
import os
from tkinter import filedialog, messagebox

def open_file(filename):
    script_dir = os.path.dirname(os.path.realpath(__file__))  
    filename = os.path.join(script_dir, filename)  
    try:
        os.startfile(filename)
    except FileNotFoundError:
        messagebox.showerror("Ошибка", f"Файл {filename} не найден")

def create_tab(tab_control, tab_name):
    tab = ttk.Frame(tab_control)
    tab_control.add(tab, text=tab_name)

    if tab_name == 'Инструкции':
        button1 = tk.Button(tab, text="Как добавить белые нули", command=lambda: open_file('инструкции/белые нули.rtf'), height=3, width=40, font=("Helvetica", 16))
        button1.pack()

        button2 = tk.Button(tab, text="Как заполнить влажность", command=lambda: open_file('инструкции/Влажность.rtf'), height=3, width=40, font=("Helvetica", 16))
        button2.pack()

        button3 = tk.Button(tab, text="Добавить грансоставы", command=lambda: open_file('инструкции/Грансостав.rtf'), height=3, width=40, font=("Helvetica", 16))
        button3.pack()

        button4 = tk.Button(tab, text="Убрать красные числа", command=lambda: open_file('инструкции/красные числа.rtf'), height=3, width=40, font=("Helvetica", 16))
        button4.pack()

        button5 = tk.Button(tab, text="Создать файлы .csv", command=lambda: open_file('инструкции/сделать csv.rtf'), height=3, width=40, font=("Helvetica", 16))
        button5.pack()
        pass

    elif tab_name == '.csv':

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
                df = pd.read_excel(file_path) 
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

        global csv_files_list
        csv_files_list = []

        global file_path
        file_path = ""

        global folder_path
        folder_path = ""

        excel_file_button = tk.Button(tab, text="Выбрать файл", command=select_excel_file, font=("Arial", 20), height=2, width=30)
        excel_file_button.pack()

        output_folder_button = tk.Button(tab, text="Выбрать папку", command=select_output_folder, font=("Arial", 20), height=2, width=30)
        output_folder_button.pack()

        convert_button = tk.Button(tab, text="Преобразовать в CSV", command=convert_to_csv, font=("Arial", 20), height=2, width=30)
        convert_button.pack()

        delete_first_line_button = tk.Button(tab, text="Удалить первую строку в файлах CSV", command=delete_first_line, font=("Arial", 20), height=2, width=30)
        delete_first_line_button.pack()

    elif tab_name == 'Влажность':
        import win32com.client
        from tkinter import Label, Entry, Button, StringVar

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

        Label(tab, text="Данные копируем из", font=("Helvetica", 16)).grid(row=0)
        Entry(tab, textvariable=range_value, font=("Helvetica", 16), width=8).grid(row=0, column=1)

        Label(tab, text="ИГЭ проверяем по столбцу", font=("Helvetica", 16)).grid(row=1)
        Entry(tab, textvariable=column_value, font=("Helvetica", 16), width=4).grid(row=1, column=1)

        Label(tab, text="ИГЭ 1", font=("Helvetica", 16)).grid(row=2)
        Entry(tab, textvariable=value1, font=("Helvetica", 16), width=4).grid(row=2, column=1)

        Label(tab, text="ИГЭ 2", font=("Helvetica", 16)).grid(row=3)
        Entry(tab, textvariable=value2, font=("Helvetica", 16), width=4).grid(row=3, column=1)

        Label(tab, text="ИГЭ 3", font=("Helvetica", 16)).grid(row=4)
        Entry(tab, textvariable=value3, font=("Helvetica", 16), width=4).grid(row=4, column=1)

        Label(tab, text="ИГЭ 4", font=("Helvetica", 16)).grid(row=5)
        Entry(tab, textvariable=value4, font=("Helvetica", 16), width=4).grid(row=5, column=1)

        Label(tab, text="Влажность проверяeм по", font=("Helvetica", 16)).grid(row=6)
        Entry(tab, textvariable=column_while, font=("Helvetica", 16), width=4).grid(row=6, column=1)

        Label(tab, text="Влажность от", font=("Helvetica", 16)).grid(row=7)
        Entry(tab, textvariable=lower_bound, font=("Helvetica", 16), width=4).grid(row=7, column=1)

        Label(tab, text="Влажность до", font=("Helvetica", 16)).grid(row=8)
        Entry(tab, textvariable=upper_bound, font=("Helvetica", 16), width=4).grid(row=8, column=1)

        Label(tab, text="Вставка со строки", font=("Helvetica", 16)).grid(row=9)
        Entry(tab, textvariable=start_value, font=("Helvetica", 16), width=4).grid(row=9, column=1)

        Label(tab, text="Столбец для вставки", font=("Helvetica", 16)).grid(row=10)
        Entry(tab, textvariable=column_select, font=("Helvetica", 16), width=4).grid(row=10, column=1)

        Button(tab, text="Обработать", command=run_macro, font=("Helvetica", 16)).grid(row=11, column=1)

        pass
#третья скрипт грансоста
    elif tab_name == 'Грансостав':
        import win32com.client
        from tkinter import Label, Entry, Button, StringVar

        def run_macro():
            g_value = g_var.get()
            s_value = s_var.get()
            m_value = m_var.get()
            values = [value.get() for value in value_vars]
            bg_value = bg_var.get()
            i_value = i_var.get()

            if g_value and s_value and m_value and all(values) and bg_value and i_value:
                xl = win32com.client.Dispatch("Excel.Application")
                xl.Visible = True

                wb = xl.ActiveWorkbook
                ws = wb.ActiveSheet

                macro_code = f"""
    Sub Kikimora()
        Dim i As Integer
        Sheets("Лист1").Select
        Dim cellValue As String
        Range("{g_value}:{s_value}").Select
        Selection.Copy
        Sheets(1).Select
        For i = {i_value} To 20000
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

                wb.VBProject.VBComponents.Remove(xlmodule)

                wb.Save()

        g_var = StringVar()
        s_var = StringVar()
        m_var = StringVar()
        value_vars = [StringVar() for _ in range(4)]
        bg_var = StringVar()
        i_var = StringVar()

        input_frame = tk.Frame(tab)
        input_frame.pack()

        file_label = Label(input_frame, text="Копировать с:", font=('Helvetica', 20))
        file_label.grid(row=0, column=0, sticky="e")
        g_entry = Entry(input_frame, textvariable=g_var, font=('Helvetica', 20), width=12)
        g_entry.grid(row=0, column=1)

        file_label = Label(input_frame, text="Копировать до:", font=('Helvetica', 20))
        file_label.grid(row=1, column=0, sticky="e")
        s_entry = Entry(input_frame, textvariable=s_var, font=('Helvetica', 20), width=12)
        s_entry.grid(row=1, column=1)

        file_label = Label(input_frame, text="Сверка ИГЭ:", font=('Helvetica', 20))
        file_label.grid(row=2, column=0, sticky="e")
        m_entry = Entry(input_frame, textvariable=m_var, font=('Helvetica', 20), width=12)
        m_entry.grid(row=2, column=1)

        for i in range(4):
            label_text = f"ИГЭ {i+1}:"
            file_label = Label(input_frame, text=label_text, font=('Helvetica', 20))
            file_label.grid(row=i+3, column=0, sticky="e")
            value_entries = Entry(input_frame, textvariable=value_vars[i], font=('Helvetica', 20), width=12)
            value_entries.grid(row=i+3, column=1)

        file_label = Label(input_frame, text="Вставка в:", font=('Helvetica', 20))
        file_label.grid(row=7, column=0, sticky="e")
        bg_entry = Entry(input_frame, textvariable=bg_var, font=('Helvetica', 20), width=12)
        bg_entry.grid(row=7, column=1)

        file_label = Label(input_frame, text="Начать со строки:", font=('Helvetica', 20))
        file_label.grid(row=8, column=0, sticky="e")
        i_entry = Entry(input_frame, textvariable=i_var, font=('Helvetica', 20), width=12)
        i_entry.grid(row=8, column=1)

        run_button = Button(tab, text="Тык! Тык!", command=run_macro, font=('Helvetica', 20))
        run_button.pack()

        pass

    elif tab_name == 'Красные числа':
        import win32com.client
        from tkinter import Label, Entry, Button, StringVar

        def run_macro():
            start_value = start_var.get()
            end_value = end_var.get()
            bf_value = bf_var.get()
            a_value = a_var.get()

            if start_value and end_value and bf_value and a_value:
                xl = win32com.client.Dispatch("Excel.Application")
                xl.Visible = True 

                wb = xl.ActiveWorkbook
                ws = wb.ActiveSheet

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

                xl.Run("CustomLoop")

                wb.VBProject.VBComponents.Remove(xlmodule)

                wb.Save()

                result_label.config(text="Макрос успешно выполнен")
            else:
                result_label.config(text="Значения для цикла не указаны")

        start_var = tk.StringVar()
        end_var = tk.StringVar()
        bf_var = tk.StringVar()
        a_var = tk.StringVar()

        start_frame = tk.Frame(tab)  
        start_frame.pack(pady=20)

        start_label = tk.Label(start_frame, text="Начало:", font=('Helvetica', 20))  
        start_label.pack(side="left")

        start_entry = tk.Entry(start_frame, textvariable=start_var, width=5, font=('Helvetica', 20))  
        start_entry.pack(side="right")

        end_frame = tk.Frame(tab)  
        end_frame.pack(pady=20)

        end_label = tk.Label(end_frame, text="Конец:", font=('Helvetica', 20))  
        end_label.pack(side="left")

        end_entry = tk.Entry(end_frame, textvariable=end_var, width=5, font=('Helvetica', 20))  
        end_entry.pack(side="right")

        bf_frame = tk.Frame(tab)  
        bf_frame.pack(pady=20)

        bf_label = tk.Label(bf_frame, text="Делать замены в столбце:", font=('Helvetica', 20))  
        bf_label.pack(side="left")

        bf_entry = tk.Entry(bf_frame, textvariable=bf_var, width=5, font=('Helvetica', 20))  
        bf_entry.pack(side="right")

        a_frame = tk.Frame(tab)  
        a_frame.pack(pady=20)

        a_label = tk.Label(a_frame, text="Красные числа в столбце:", font=('Helvetica', 20))  
        a_label.pack(side="left")

        a_entry = tk.Entry(a_frame, textvariable=a_var, width=5, font=('Helvetica', 20))  
        a_entry.pack(side="right")

        run_button = tk.Button(tab, text="Убрать красное", command=run_macro, font=('Helvetica', 20))  
        run_button.pack(pady=20)

        result_label = tk.Label(tab, text="", font=('Helvetica', 20))  
        result_label.pack()

        pass

    elif tab_name == 'Белые нули':
        import win32com.client
        from tkinter import Label, Entry, Button, StringVar

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

        v_start_label = tk.Label(tab, text="Проверка от:", font=('Helvetica', 30))
        v_start_entry = tk.Entry(tab, width=7, font=('Helvetica', 30))
        v_end_label = tk.Label(tab, text="Проверка до:", font=('Helvetica', 30))
        v_end_entry = tk.Entry(tab, width=7, font=('Helvetica', 30)) 
        k_label = tk.Label(tab, text="Вставка от:", font=('Helvetica', 30))
        k_entry = tk.Entry(tab, width=7, font=('Helvetica', 30))
        u_label = tk.Label(tab, text="Вставка до:", font=('Helvetica', 30))
        u_entry = tk.Entry(tab, width=7, font=('Helvetica', 30))
        button = tk.Button(tab, text="Напихать нулей", command=generate_macro, font=('Helvetica', 20))

        v_start_label.grid(row=0, column=0)
        v_start_entry.grid(row=0, column=1)
        v_end_label.grid(row=1, column=0)
        v_end_entry.grid(row=1, column=1)
        k_label.grid(row=2, column=0)
        k_entry.grid(row=2, column=1)
        u_label.grid(row=3, column=0)
        u_entry.grid(row=3, column=1)
        button.grid(row=4, column=0, columnspan=2)
        
        pass

    elif tab_name == 'Кнопка':
        button1 = tk.Button(tab, text="Кнопка любви к Дыне и Гренке", height=15, width=40, font=("Helvetica", 16))
        button1.pack()

        pass

root = tk.Tk()

tabControl = ttk.Notebook(root)

tab1 = create_tab(tabControl, 'Инструкции')
tab2 = create_tab(tabControl, 'Влажность')
tab3 = create_tab(tabControl, 'Грансостав')
tab4 = create_tab(tabControl, 'Красные числа')
tab5 = create_tab(tabControl, 'Белые нули')
tab6 = create_tab(tabControl, '.csv')
tab7 = create_tab(tabControl, 'Кнопка')

tabControl.pack(expand=1, fill="both")

root.mainloop()
