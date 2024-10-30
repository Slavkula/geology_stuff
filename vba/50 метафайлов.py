import win32com.client
from tkinter import Tk
from tkinter.filedialog import askopenfilename

def run_excel_macro():
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True

    # Open file dialog to select Excel file
    root = Tk()
    root.withdraw()  # Hide the root window
    file_path = askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    root.destroy()  # Close the root window

    if not file_path:
        print("No file selected")
        return

    workbook = excel.Workbooks.Open(file_path)

    # Insert your VBA macro
    macro_code = """
    Sub ExcelWordMetaPaste()
        Dim WordApp As Object
        On Error Resume Next
        Set WordApp = GetObject(, "Word.Application")
        On Error GoTo 0
        If WordApp Is Nothing Then
            Set WordApp = CreateObject("Word.Application")
        End If
        If WordApp.Documents.Count = 0 Then
            MsgBox "Нет открытых документов в Word", vbExclamation
            Exit Sub
        End If
        Dim WordDoc As Object
        Set WordDoc = WordApp.ActiveDocument
        
        Dim Ranges As Variant
        Ranges = Array("B1:D22", "G1:I22", "L1:N22", "R1:T22", "W1:Y22", "AB1:AD22", "AI1:AK22", "AN1:AP22", "AS1:AU22", "B23:D44", "G23:I44", "L23:N44", "R23:T44", "W23:Y44", "AB23:AD44", "AI23:AK44", "AN23:AP44", "AS23:AU44", "B45:D134", "G45:I134", "L45:N134", "R45:T134", "W45:Y134", "AB45:AD134", "AI45:AK134", "AN45:AP134", "AS45:AU134", "B135:D157", "G135:I157", "L135:N157", "R135:T157", "W135:Y157", "AB135:AD157", "AI135:AK157", "AN135:AP157", "AS135:AU157", "B159:D180", "G159:I180", "L159:N180", "R159:T180", "W159:Y180", "AB159:AD180", "AI159:AK180", "AN159:AP180", "AS159:AU180", "B183:D204", "G183:I204", "L183:N204", "R183:T204", "W183:Y204", "AB183:AD204", "AI183:AK204", "AN183:AP204", "AS183:AU204")
        Dim rng As Variant
        For Each rng In Ranges
            Range(rng).Copy
            WordApp.Selection.PasteSpecial Link:=False, DataType:=wdPasteEnhancedMetafile, Placement:=wdInLine, DisplayAsIcon:=False
            WordApp.Selection.TypeParagraph
            WordApp.Selection.TypeParagraph
        Next rng
        Application.CutCopyMode = False
        Set WordDoc = Nothing
        Set WordApp = Nothing
    End Sub
    """
    module = workbook.VBProject.VBComponents.Add(1)  # 1 means vbext_ct_StdModule
    module.CodeModule.AddFromString(macro_code)

    # Run the macro
    excel.Application.Run("ExcelWordMetaPaste")

    # Cleanup
    workbook.Close(SaveChanges=False)
    excel.Quit()

if __name__ == "__main__":
    run_excel_macro()
