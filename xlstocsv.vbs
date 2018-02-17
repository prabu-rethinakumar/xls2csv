if WScript.Arguments.Count < 2 Then
    WScript.Echo "Kindly specify the source path for xls and the destination for csv. Usage: xls2csv source_path/excel.xls target_dir/csvfile.csv"
    Wscript.Quit
End If
Dim oExcel
Set oExcel = CreateObject("Excel.Application")
Dim oBook
Set oBook = oExcel.Workbooks.Open(Wscript.Arguments.Item(0))
oBook.SaveAs WScript.Arguments.Item(1), 6
oBook.Close False
oExcel.Quit
WScript.Echo "Done"
