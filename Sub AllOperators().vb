Sub DeleteRows()
    Dim lastRow As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Requisitions")

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow > 2 Then
        ws.Rows("3:" & lastRow).Delete
    End If
End Sub

Sub DeleteRows1()
    Dim lastRow As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Hires")

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow > 2 Then
        ws.Rows("3:" & lastRow).Delete
    End If
End Sub

Sub DeleteRows2()
    Dim lastRow As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Applications")

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow > 2 Then
        ws.Rows("3:" & lastRow).Delete
    End If
End Sub
