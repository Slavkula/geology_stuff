import tkinter as tk
from tkinter import ttk

def create_tab(tab_control, tab_name):
    tab = ttk.Frame(tab_control)
    tab_control.add(tab, text=tab_name)

    if tab_name == 'Вкладка 1':
        #кукуруза
        pass

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
#третья скрипт
    elif tab_name == 'Грансостав':
        import win32com.client
        from tkinter import Label, Entry, Button, StringVar

        def run_macro():
            g_value = g_var.get()
            s_value = s_var.get()
            m_value = m_var.get()
            values = [value.get() for value in value_vars if value.get()]
            bg_value = bg_var.get()
            i_value = i_var.get()

            if g_value and s_value and m_value and bg_value and i_value:
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
            If { ' Or '.join(f'cellValue = "{value}"' for value in values) } Then
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



root = tk.Tk()

tabControl = ttk.Notebook(root)

tab1 = create_tab(tabControl, 'Вкладка 1')
tab2 = create_tab(tabControl, 'Влажность')
tab3 = create_tab(tabControl, 'Грансостав')
tab4 = create_tab(tabControl, 'Вкладка 4')
tab5 = create_tab(tabControl, 'Вкладка 5')
tab6 = create_tab(tabControl, 'Вкладка 6')
tab7 = create_tab(tabControl, 'Вкладка 7')

tabControl.pack(expand=1, fill="both")

root.mainloop()
