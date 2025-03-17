Sub AddNewColumns()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Data Integrity - Positions Copy")  
    
    Dim i As Integer
    For i = 1 To 7
        ws.Columns("AE:AE").Insert Shift:=xlToRight
    Next i
    
    ws.Range("AE1").Value = "CCID"
    ws.Range("AF1").Value = "Fin 4"
    ws.Range("AG1").Value = "Fin 2"
    ws.Range("AH1").Value = "Fin 3"
    ws.Range("AI1").Value = "Fin 5 - Business Line"
    ws.Range("AJ1").Value = "Fin 7 - Sub-branch"
    ws.Range("AK1").Value = "GEO"
    ws.Range("AL1").Value = "Exclude"
    
    ws.Range("AD:AD").Copy
    ws.Range("AE:AL").PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

    ws.Range("AE2").Formula = "=IF(IFERROR(IFERROR(VALUE(LEFT(IF(LEFT(RIGHT(AA2,10),1)=""("",RIGHT(AA2,9),RIGHT(AA2,11)),LEN(IF(LEFT(RIGHT(AA2,10),1)=""("",RIGHT(AA2,9),RIGHT(AA2,11)))-1)),LEFT(IF(LEFT(RIGHT(AA2,10),1)=""("",RIGHT(AA2,9),RIGHT(AA2,11)),LEN(IF(LEFT(RIGHT(AA2,10),1)=""("",RIGHT(AA2,9),RIGHT(AA2,11)))-1)),AB2)=0,""Unmapped"",IFERROR(IFERROR(VALUE(LEFT(IF(LEFT(RIGHT(AA2,10),1)=""("",RIGHT(AA2,9),RIGHT(AA2,11)),LEN(IF(LEFT(RIGHT(AA2,10),1)=""("",RIGHT(AA2,9),RIGHT(AA2,11)))-1)),LEFT(IF(LEFT(RIGHT(AA2,10),1)=""("",RIGHT(AA2,9),RIGHT(AA2,11)),LEN(IF(LEFT(RIGHT(AA2,10),1)=""("",RIGHT(AA2,9),RIGHT(AA2,11)))-1)),AB2))"

ws.Range("AE2:AE" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row).FillDown

ws.Range("AE:AE").Value = ws.Range("AE:AE").Value

    With ws
        .Columns("B").TextToColumns Destination:=.Range("B1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), _
        TrailingMinusNumbers:=True
        .Columns("AC").TextToColumns Destination:=.Range("AC1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), _
        TrailingMinusNumbers:=True
        .Columns("AD").TextToColumns Destination:=.Range("AD1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), _
        TrailingMinusNumbers:=True

        .Columns("M").Replace What:=" (*)", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        .Columns("N").Replace What:=" (*)", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        
        .Columns("H").Replace What:="??-", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    
    lastRow = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row
    
    ws.Range("E2:E" & lastRow).NumberFormat = "m/d/yyyy"
        End With

End Sub

Sub AddVLOOKUPAndCopyAsValues()

    Dim LastRow As Long

    With ThisWorkbook.Sheets("Data Integrity - Positions Copy") 
        .Range("AF2").Formula = "=IFNA(XLOOKUP(ae2,[CC_MAP.xlsx]Sheet1!$A:$A,[CC_MAP.xlsx]Sheet1!$G:$G), IF(LEFT(N2, 1) = ""A"", ""Atos"", ""Eviden""))"
        .Range("AG2").Formula = "=IFNA(VLOOKUP(AE2,[CC_MAP.xlsx]Sheet1!$A:$E,5,0),XLOOKUP(AH2,[CC_MAP.xlsx]Sheet1!$F:$F,[CC_MAP.xlsx]Sheet1!$E:$E))"
        .Range("AH2").Formula = "=IFNA(VLOOKUP(AE2,[CC_MAP.xlsx]Sheet1!$A:$F,6,0),H2)"
        .Range("AI2").Formula = "=IFNA(VLOOKUP(AE2,[CC_MAP.xlsx]Sheet1!$A:$H,8,0),XLOOKUP(AH2,[CC_MAP.xlsx]Sheet1!$F:$F,[CC_MAP.xlsx]Sheet1!$H:$H))"
        .Range("AJ2").Formula = "=IFNA(VLOOKUP(AE2,[CC_MAP.xlsx]Sheet1!$A:$J,10,0),XLOOKUP(AH2,[CC_MAP.xlsx]Sheet1!$F:$F,[CC_MAP.xlsx]Sheet1!$J:$J))"
        .Range("AK2").Formula = "=IFNA(VLOOKUP(AE2,[CC_MAP.xlsx]Sheet1!$A:$M,4,0),XLOOKUP(ag2,[CC_MAP.xlsx]Sheet1!$e:$e,[CC_MAP.xlsx]Sheet1!$d:$d))"
        .Range("AL2").Formula = "=IF(OR(" & _
            "A2='[PL de sters DI 23.07.2024 1.xlsx]Sheet1'!$A$2," & _
            "A2='[PL de sters DI 23.07.2024 1.xlsx]Sheet1'!$A$3," & _
            "A2='[PL de sters DI 23.07.2024 1.xlsx]Sheet1'!$A$4," & _
            "A2='[PL de sters DI 23.07.2024 1.xlsx]Sheet1'!$A$5," & _
            "A2='[PL de sters DI 23.07.2024 1.xlsx]Sheet1'!$A$6," & _
            "A2='[PL de sters DI 23.07.2024 1.xlsx]Sheet1'!$A$7," & _
            "A2='[PL de sters DI 23.07.2024 1.xlsx]Sheet1'!$A$8," & _
            "A2='[PL de sters DI 23.07.2024 1.xlsx]Sheet1'!$A$9," & _
            "A2='[PL de sters DI 23.07.2024 1.xlsx]Sheet1'!$A$10," & _
            "A2='[PL de sters DI 23.07.2024 1.xlsx]Sheet1'!$A$11)," & _
            """exclude"",""OK"")"

        LastRow = .Cells(.Rows.Count, "AE").End(xlUp).Row

        .Range("AF2:AF" & LastRow).FillDown
        .Range("AG2:AG" & LastRow).FillDown
        .Range("AH2:AH" & LastRow).FillDown
        .Range("AI2:AI" & LastRow).FillDown
        .Range("AJ2:AJ" & LastRow).FillDown
        .Range("AK2:AK" & LastRow).FillDown
        .Range("AL2:AL" & LastRow).FillDown

        .Range("AF2:AL" & LastRow).Copy
        .Range("AF2:AL" & LastRow).PasteSpecial xlPasteValues
        Application.CutCopyMode = False
    End With

End Sub