Sub AddNewColumns()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("KPI Candidate Progress 2024 w C")

    Dim i As Integer
    For i = 1 To 7
        ws.Columns("AC:AC").Insert Shift:=xlToRight
    Next i

    ws.Range("AC1").Value = "CCID"
    ws.Range("AD1").Value = "FIN 4"
    ws.Range("AE1").Value = "Fin 2"
    ws.Range("AF1").Value = "Fin 3"
    ws.Range("AG1").Value = "Fin 5"
    ws.Range("AH1").Value = "Fin 7"
    ws.Range("AI1").Value = "GEO"

    ws.Range("A:A").Copy
    ws.Range("AC:AI").PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

                ws.Range("AC2").Formula = "=IF(IFERROR(IFERROR(VALUE(LEFT(IF(LEFT(RIGHT(AA2,10),1)=""("",RIGHT(AA2,9),RIGHT(AA2,11)),LEN(IF(LEFT(RIGHT(AA2,10),1)=""("",RIGHT(AA2,9),RIGHT(AA2,11)))-1)),LEFT(IF(LEFT(RIGHT(AA2,10),1)=""("",RIGHT(AA2,9),RIGHT(AA2,11)),LEN(IF(LEFT(RIGHT(AA2,10),1)=""("",RIGHT(AA2,9),RIGHT(AA2,11)))-1)),AB2)=0,""Unmapped"",IFERROR(IFERROR(VALUE(LEFT(IF(LEFT(RIGHT(AA2,10),1)=""("",RIGHT(AA2,9),RIGHT(AA2,11)),LEN(IF(LEFT(RIGHT(AA2,10),1)=""("",RIGHT(AA2,9),RIGHT(AA2,11)))-1)),LEFT(IF(LEFT(RIGHT(AA2,10),1)=""("",RIGHT(AA2,9),RIGHT(AA2,11)),LEN(IF(LEFT(RIGHT(AA2,10),1)=""("",RIGHT(AA2,9),RIGHT(AA2,11)))-1)),AB2))"
ws.Range("AC2:AC" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row).FillDown

ws.Range("AC:AC").Value = ws.Range("AC:AC").Value
End Sub

Sub ApplyTextToColumns()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("KPI Candidate Progress 2024 w C")

     With ws
        .Columns("B").TextToColumns Destination:=.Range("B1"), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
        .Columns("E").TextToColumns Destination:=.Range("E1"), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
        .Columns("X").TextToColumns Destination:=.Range("X1"), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
        .Columns("M").TextToColumns Destination:=.Range("M1"), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
        .Columns("P").TextToColumns Destination:=.Range("P1"), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True

        .Columns("H").Replace What:=" (*)", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        .Columns("G").Replace What:="??-", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

        .Range("AD2").Formula = "=IFNA(VLOOKUP(AC2, [CC_MAP.xlsx]Sheet1!$A:$G,7,0),IF(LEFT(H2,1)=""A"",""Atos"",""Eviden""))"
        .Range("AE2").Formula = "=IFNA(VLOOKUP(AC2, [CC_MAP.xlsx]Sheet1!$A:$E,5,0),XLOOKUP(af2,[CC_MAP.xlsx]Sheet1!$F:$F,[CC_MAP.xlsx]Sheet1!$E:$E))"
        .Range("AF2").Formula = "=IFNA(VLOOKUP(AC2, [CC_MAP.xlsx]Sheet1!$A:$F,6,0),G2)"
        .Range("AG2").Formula = "=IFNA(VLOOKUP(AC2, [CC_MAP.xlsx]Sheet1!$A:$H,8,0),XLOOKUP(af2,[CC_MAP.xlsx]Sheet1!$F:$F,[CC_MAP.xlsx]Sheet1!$H:$H))"
        .Range("AH2").Formula = "=IFNA(VLOOKUP(AC2, [CC_MAP.xlsx]Sheet1!$A:$J,10,0),XLOOKUP(af2,[CC_MAP.xlsx]Sheet1!$F:$F,[CC_MAP.xlsx]Sheet1!$J:$J))"
        .Range("AI2").Formula = "=IFNA(XLOOKUP(ac2,[CC_MAP.xlsx]Sheet1!$A:$A,[CC_MAP.xlsx]Sheet1!$D:$D),XLOOKUP(af2,[CC_MAP.xlsx]Sheet1!$F:$F,[CC_MAP.xlsx]Sheet1!$D:$D))"

    ws.Range("AD2:AD" & ws.Cells(ws.Rows.Count, "AC").End(xlUp).Row).FillDown
    ws.Range("AE2:AE" & ws.Cells(ws.Rows.Count, "AC").End(xlUp).Row).FillDown
    ws.Range("AF2:AF" & ws.Cells(ws.Rows.Count, "AC").End(xlUp).Row).FillDown
    ws.Range("AG2:AG" & ws.Cells(ws.Rows.Count, "AC").End(xlUp).Row).FillDown
    ws.Range("AH2:AH" & ws.Cells(ws.Rows.Count, "AC").End(xlUp).Row).FillDown
    ws.Range("AI2:AI" & ws.Cells(ws.Rows.Count, "AC").End(xlUp).Row).FillDown

   ws.Range("AD2:AI" & ws.Cells(ws.Rows.Count, "AC").End(xlUp).Row).FillDown
    End With
End Sub

Sub CreatePivotTable()
    Dim PSheet As Worksheet
    Dim DSheet As Worksheet
    Dim PCache As PivotCache
    Dim PTable As PivotTable
    Dim PRange As Range
    Dim LastRow As Long
    Dim LastCol As Long

    Set DSheet = Worksheets("KPI Candidate Progress 2024 w C")

    LastRow = DSheet.Cells(DSheet.Rows.Count, 1).End(xlUp).Row
    LastCol = DSheet.Cells(1, DSheet.Columns.Count).End(xlToLeft).Column
    Set PRange = DSheet.Cells(1, 1).Resize(LastRow, LastCol)

    Set PCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=PRange)

    Set PSheet = Worksheets.Add

    Set PTable = PCache.CreatePivotTable( _
        TableDestination:=PSheet.Cells(1, 1), _
        TableName:="PivotTable1")

    With PTable
        .PivotFields("Fin 4").Orientation = xlRowField
        With .PivotFields("Job Req ID")
            .Orientation = xlDataField
            .Function = xlCount
            .Name = "Count of Job Req ID"
        End With
    End With
End Sub

