Option Explicit

Dim fso

systemutil.CloseProcessByName("Excel.exe")

Set fso = createobject("Scripting.FileSystemObject")

ReadExcelData()
