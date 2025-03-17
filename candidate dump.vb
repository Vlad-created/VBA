Sub ApplyTextToColumns()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Base")

    With ws
        .Columns("B").TextToColumns Destination:=.Range("B1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), _
        TrailingMinusNumbers:=True
        
        .Columns("C").TextToColumns Destination:=.Range("C1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), _
        TrailingMinusNumbers:=True
        
        .Columns("K").TextToColumns Destination:=.Range("K1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), _
        TrailingMinusNumbers:=True
        
        .Columns("L").TextToColumns Destination:=.Range("L1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), _
        TrailingMinusNumbers:=True
        
        .Columns("T").TextToColumns Destination:=.Range("T1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), _
        TrailingMinusNumbers:=True
        
        .Columns("W").TextToColumns Destination:=.Range("W1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), _
        TrailingMinusNumbers:=True
        
        .Columns("Z").TextToColumns Destination:=.Range("Z1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), _
        TrailingMinusNumbers:=True
        
        .Columns("AR").TextToColumns Destination:=.Range("AR1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), _
        TrailingMinusNumbers:=True
        
        .Columns("AP").Replace What:=" (*)", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        .Columns("AQ").Replace What:=" (*)", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

        .Columns("AI").Replace What:="??-", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

        lastRow = ws.Cells(ws.Rows.Count, "AZ").End(xlUp).Row

' Drag formula down
        ws.Range("BA2:CK2").AutoFill Destination:=ws.Range("BA2:CK" & lastRow), Type:=xlFillDefault

' Copy and paste as values from BA3 to CK3 to the last row
        ws.Range("BA3:CK3").AutoFill Destination:=ws.Range("BA3:CK" & lastRow), Type:=xlFillDefault
        ws.Range("BA3:CK" & lastRow).Copy
        ws.Range("BA3:CK" & lastRow).PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False

    End With
End Sub

Sub DeleteRows()
    Dim lastRow As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Base")

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow > 2 Then
        ws.Rows("3:" & lastRow).Delete
    End If
End Sub
