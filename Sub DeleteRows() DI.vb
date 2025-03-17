Sub DeleteRows()
    Dim lastRow As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Positions") 

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow > 2 Then
        ws.Rows("3:" & lastRow).Delete
    End If
End Sub

Sub DeleteRows1()
    Dim lastRow As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Applications") 

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow > 2 Then
        ws.Rows("3:" & lastRow).Delete
    End If
End Sub