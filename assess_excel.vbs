Dim args, objExcel

set args = WScript.Arguments
set objExcel = CreateObject("Excel.Application")


objExcel.Workbooks.Open args(0)
'objExcel.Workbooks.Open "<file_path>"
objExcel.Visible = False

    objExcel.Run "<macro_name>" ' the macro should put in module 

objExcel.ActiveWorkbook.Save
objExcel.ActiveWorkbook.Close(0)
objExcel.Quit
