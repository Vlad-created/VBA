Sub AddNewColumns()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("GOA Report Copy(2) Copy(2) Copy")

    Dim i As Integer
    For i = 1 To 2
        ws.Columns("O:O").Insert Shift:=xlToRight
    Next i

    ws.Range("O1").Value = "CCID"
    ws.Range("P1").Value = "HM"

    ws.Range("N:N").Copy
    ws.Range("O:P").PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

    ws.Range("O2").Formula = "=IF(IFERROR(IFERROR(VALUE(LEFT(IF(LEFT(RIGHT(K2,10),1)=""("",RIGHT(K2,9),RIGHT(K2,11)),LEN(IF(LEFT(RIGHT(K2,10),1)=""("",RIGHT(K2,9),RIGHT(K2,11)))-1)),LEFT(IF(LEFT(RIGHT(K2,10),1)=""("",RIGHT(K2,9),RIGHT(K2,11)),LEN(IF(LEFT(RIGHT(K2,10),1)=""("",RIGHT(K2,9),RIGHT(K2,11)))-1)),L2)=0,""Unmapped"",IFERROR(IFERROR(VALUE(LEFT(IF(LEFT(RIGHT(K2,10),1)=""("",RIGHT(K2,9),RIGHT(K2,11)),LEN(IF(LEFT(RIGHT(K2,10),1)=""("",RIGHT(K2,9),RIGHT(K2,11)))-1)),LEFT(IF(LEFT(RIGHT(K2,10),1)=""("",RIGHT(K2,9),RIGHT(K2,11)),LEN(IF(LEFT(RIGHT(K2,10),1)=""("",RIGHT(K2,9),RIGHT(K2,11)))-1)),L2))"
    
    ws.Range("O2:O" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row).FillDown

    ws.Range("O:O").Value = ws.Range("O:O").Value

    ws.Range("P2").Formula = "=CONCAT(C2, "" "", D2)"

    ws.Range("P2:P" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row).FillDown

    ws.Range("P:P").Value = ws.Range("P:P").Value

    With ws
        .Columns("B").TextToColumns Destination:=.Range("B1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        SemiColon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), _
        TrailingMinusNumbers:=True
        .Columns("F").TextToColumns Destination:=.Range("F1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        SemiColon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), _
        TrailingMinusNumbers:=True
        .Columns("J").TextToColumns Destination:=.Range("J1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        SemiColon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), _
        TrailingMinusNumbers:=True
    End With

    ws.Columns("M").Replace What:=" (*)", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    ws.Columns("N").Replace What:=" (*)", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

    ws.Columns("E").NumberFormat = "m/d/yyyy"

    ws.Columns("C:D").Delete

End Sub