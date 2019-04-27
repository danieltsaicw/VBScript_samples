Dim args, objExcel

set args = WScript.Arguments
set objExcel = CreateObject("Excel.Application")

'Safe Mode
'Set objWshShell = CreateObject("WScript.Shell")
'objWshShell.Run "Excel.exe /s"


objExcel.Workbooks.Open args(0)
objExcel.Visible = False

objExcel.Run "main"

objExcel.ActiveWorkbook.Save
objExcel.ActiveWorkbook.Close(0)
objExcel.Quit
