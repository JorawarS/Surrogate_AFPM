Sub combinations()
    Dim wrksht As Worksheet
    Dim gap As Double
    Dim mat As String
    Dim pm_thic As Double
    Dim pm_angle As Double
    Set wrksht = ActiveWorkbook.Worksheets("Combinations")
    Application.DisplayAlerts = False
    ' Select cell A2, *first line of data*.
      wrksht.Range("A10700").Select
      ' Set Do loop to stop when an empty cell is reached.
      Do Until IsEmpty(ActiveCell)
         ' Extract Parameters from Current Row
         gap = ActiveCell.Value
         mat = ActiveCell.Offset(0, 1).Value
         pm_thic = ActiveCell.Offset(0, 2).Value
         pm_angle = ActiveCell.Offset(0, 3).Value
         'MsgBox gap & vbNewLine & mat & vbNewLine & pm_thic & vbNewLine & pm_angle
         ' Call MAGNET to solve using given parameters
         Call square(gap, mat, pm_thic, pm_angle)
         Call import_results
         ' Step down 1 row from present location.
         ActiveCell.Offset(1, -8).Select
      Loop

End Sub
