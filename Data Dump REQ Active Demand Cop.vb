Sub AddNewColumns()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Data Dump REQ Active Demand")
    
    Dim i As Integer
    For i = 1 To 11
        ws.Columns("AY:AY").Insert Shift:=xlToRight
    Next i
    
    ws.Range("AY1").Value = "CCID"
    ws.Range("AZ1").Value = "Fin 4"
    ws.Range("BA1").Value = "Fin 2"
    ws.Range("BB1").Value = "Fin 3"
    ws.Range("BC1").Value = "Fin 5 - Business Line"
    ws.Range("BD1").Value = "Fin 7 - Sub-branch"
    ws.Range("BE1").Value = "GEO"
    ws.Range("BF1").Value = "rec"
    ws.Range("BG1").Value = "supp"
    ws.Range("BH1").Value = "supp team"
    ws.Range("BI1").Value = "rec group"

    
    ws.Range("AX:AX").Copy
    ws.Range("AY:BI").PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

    ws.Range("AY2").Formula = "=IF(IFERROR(IFERROR(VALUE(LEFT(IF(LEFT(RIGHT(AC2,10),1)=""("",RIGHT(AC2,9),RIGHT(AC2,11)),LEN(IF(LEFT(RIGHT(AC2,10),1)=""("",RIGHT(AC2,9),RIGHT(AC2,11)))-1)),LEFT(IF(LEFT(RIGHT(AC2,10),1)=""("",RIGHT(AC2,9),RIGHT(AC2,11)),LEN(IF(LEFT(RIGHT(AC2,10),1)=""("",RIGHT(AC2,9),RIGHT(AC2,11)))-1)),AB2)=0,""Unmapped"",IFERROR(IFERROR(VALUE(LEFT(IF(LEFT(RIGHT(AC2,10),1)=""("",RIGHT(AC2,9),RIGHT(AC2,11)),LEN(IF(LEFT(RIGHT(AC2,10),1)=""("",RIGHT(AC2,9),RIGHT(AC2,11)))-1)),LEFT(IF(LEFT(RIGHT(AC2,10),1)=""("",RIGHT(AC2,9),RIGHT(AC2,11)),LEN(IF(LEFT(RIGHT(AC2,10),1)=""("",RIGHT(AC2,9),RIGHT(AC2,11)))-1)),AB2))"
ws.Range("AY2:AY" & ws.Cells(ws.Rows.Count, "AL").End(xlUp).Row).FillDown

ws.Range("AY:AY").Value = ws.Range("AY:AY").Value
End Sub

Sub ApplyTextToColumns()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Data Dump REQ Active Demand")
    
    With ws
        .Columns("B").TextToColumns Destination:=.Range("B1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), _
        TrailingMinusNumbers:=True
        
        .Columns("C").TextToColumns Destination:=.Range("C1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), _
        TrailingMinusNumbers:=True
        
        .Columns("O").TextToColumns Destination:=.Range("O1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), _
        TrailingMinusNumbers:=True
        
        .Columns("AJ").TextToColumns Destination:=.Range("AJ1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), _
        TrailingMinusNumbers:=True
        
        .Columns("AM").TextToColumns Destination:=.Range("AM1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), _
        TrailingMinusNumbers:=True
        .Columns("J").Replace What:=" (*)", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        .Columns("K").Replace What:=" (*)", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        
        .Columns("W").Replace What:="??-", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

.Range("AZ2").Formula = "=IFNA(XLOOKUP(AY2, [CC_MAP.xlsx]Sheet1!$A:$A, [CC_MAP.xlsx]Sheet1!$G:$G), IF(LEFT(J2, 1) = ""A"", ""Atos"", ""Eviden""))"
.Range("BA2").Formula = "=IFNA(VLOOKUP(AY2,[CC_MAP.xlsx]Sheet1!$A:$E,5,0),XLOOKUP(BB2,[CC_MAP.xlsx]Sheet1!$F:$F,[CC_MAP.xlsx]Sheet1!$E:$E))"
.Range("BB2").Formula = "=IFNA(VLOOKUP(AY2,[CC_MAP.xlsx]Sheet1!$A:$F,6,0),W2)"
.Range("BC2").Formula = "=IFNA(VLOOKUP(AY2,[CC_MAP.xlsx]Sheet1!$A:$H,8,0),XLOOKUP(BB2,[CC_MAP.xlsx]Sheet1!$F:$F,[CC_MAP.xlsx]Sheet1!$H:$H))"
.Range("BD2").Formula = "=IFNA(VLOOKUP(AY2,[CC_MAP.xlsx]Sheet1!$A:$J,10,0),XLOOKUP(BB2,[CC_MAP.xlsx]Sheet1!$F:$F,[CC_MAP.xlsx]Sheet1!$J:$J))"
.Range("BE2").Formula = "=IFNA(VLOOKUP(AY2,[CC_MAP.xlsx]Sheet1!$A:$M,4,0),XLOOKUP(ba2,[CC_MAP.xlsx]Sheet1!$e:$e,[CC_MAP.xlsx]Sheet1!$d:$d))"
.Range("BF2").Formula = "=IFNA(VLOOKUP(AJ2,'[MasterList_with CC.xlsx]MasterList'!$D:$AJ,19,FALSE),IFNA(VLOOKUP(AJ2,'[MasterList_with CC.xlsx]MasterList'!$C:$AJ,20,FALSE),IFNA(VLOOKUP(AJ2,'[MasterList_with CC.xlsx]MasterList'!$B:$AJ,21,FALSE),VLOOKUP(AJ2,'[MasterList_with CC.xlsx]MasterList'!$D:$AJ,19,FALSE))))"
.Range("BG2").Formula = "=IFNA(VLOOKUP(AM2,'[MasterList_with CC.xlsx]MasterList'!$D:$AJ,19,FALSE),IFNA(VLOOKUP(AM2,'[MasterList_with CC.xlsx]MasterList'!$C:$AJ,20,FALSE),IFNA(VLOOKUP(AM2,'[MasterList_with CC.xlsx]MasterList'!$B:$AJ,21,FALSE),VLOOKUP(AM2,'[MasterList_with CC.xlsx]MasterList'!$D:$AJ,19,FALSE))))"
.Range("BH2").Formula = "=IFNA(VLOOKUP(AM2,'[MasterList_with CC.xlsx]MasterList'!$D:$AJ,5,FALSE),IFNA(VLOOKUP(AM2,'[MasterList_with CC.xlsx]MasterList'!$C:$AJ,6,FALSE),IFNA(VLOOKUP(AM2,'[MasterList_with CC.xlsx]MasterList'!$B:$AJ,7,FALSE),VLOOKUP(AM2,'[MasterList_with CC.xlsx]MasterList'!$D:$AJ,5,FALSE))))"
.Range("BI2").Formula = "=IFNA(VLOOKUP(AJ2,'[MasterList_with CC.xlsx]MasterList'!$D:$AJ,17,FALSE),IFNA(VLOOKUP(AJ2,'[MasterList_with CC.xlsx]MasterList'!$C:$AJ,18,FALSE),IFNA(VLOOKUP(AJ2,'[MasterList_with CC.xlsx]MasterList'!$B:$AJ,19,FALSE),VLOOKUP(AJ2,'[MasterList_with CC.xlsx]MasterList'!$D:$AJ,17,FALSE))))"


ws.Range("AZ2:AZ" & ws.Cells(ws.Rows.Count, "AY").End(xlUp).Row).FillDown
ws.Range("BA2:BA" & ws.Cells(ws.Rows.Count, "AY").End(xlUp).Row).FillDown
ws.Range("BB2:BB" & ws.Cells(ws.Rows.Count, "AY").End(xlUp).Row).FillDown
ws.Range("BC2:BC" & ws.Cells(ws.Rows.Count, "AY").End(xlUp).Row).FillDown
ws.Range("BD2:BD" & ws.Cells(ws.Rows.Count, "AY").End(xlUp).Row).FillDown
ws.Range("BE2:BE" & ws.Cells(ws.Rows.Count, "AY").End(xlUp).Row).FillDown
ws.Range("BF2:BF" & ws.Cells(ws.Rows.Count, "AY").End(xlUp).Row).FillDown
ws.Range("BG2:BG" & ws.Cells(ws.Rows.Count, "AY").End(xlUp).Row).FillDown
ws.Range("BH2:BH" & ws.Cells(ws.Rows.Count, "AY").End(xlUp).Row).FillDown
ws.Range("BI2:BI" & ws.Cells(ws.Rows.Count, "AY").End(xlUp).Row).FillDown

ws.Range("AZ2:BI" & ws.Cells(ws.Rows.Count, "AY").End(xlUp).Row).Value = ws.Range("AZ2:BI" & ws.Cells(ws.Rows.Count, "AY").End(xlUp).Row).Value

.Columns("AC:AC").Delete
End With
End Sub

