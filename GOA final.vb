Sub CopyFormulasDown()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim col As Long

    Set ws = ThisWorkbook.Sheets("Offer Requests")

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    For col = 2 To 52
        ws.Range(ws.Cells(2, col), ws.Cells(lastRow, col)).Formula = ws.Cells(2, col).Formula
    Next col
End Sub 