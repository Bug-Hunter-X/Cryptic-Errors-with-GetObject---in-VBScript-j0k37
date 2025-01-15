On Error Resume Next

Dim objExcel

' **Improved path handling and error checking**
Set objExcel = GetObject("C:\\path\\to\\excel.exe", "Excel.Application")

If Err.Number <> 0 Then
  MsgBox "Error accessing Excel: " & Err.Number & " - " & Err.Description, vbCritical
  Err.Clear
  ' Handle the error appropriately - e.g., use a different method or exit gracefully
  WScript.Quit
Else
  ' Work with the Excel object
  objExcel.Visible = True
  objExcel.Workbooks.Add
  ' ... Rest of your Excel automation code ...
  objExcel.Quit
  Set objExcel = Nothing
End If

On Error GoTo 0