Sub import_results()
Workbooks.Open Filename:="Z:/body force.csv"
'Sheets("body force").Copy After:=Workbooks(ThisWorkbook.Name).Sheets(2)
Workbooks(2).Activate
ActiveSheet.Range("E26:E29").Copy
Workbooks(ThisWorkbook.Name).Sheets(2).Activate
ActiveSheet.Range(ActiveCell.Offset(0, 4), ActiveCell.Offset(0, 7)).PasteSpecial Transpose:=True
Workbooks(2).Activate
ActiveSheet.Range("I26:I29").Copy
Workbooks(ThisWorkbook.Name).Sheets(2).Activate
ActiveSheet.Range(ActiveCell.Offset(0, 4), ActiveCell.Offset(0, 7)).PasteSpecial Transpose:=True
Workbooks(2).Close SaveChanges:=False

End Sub
