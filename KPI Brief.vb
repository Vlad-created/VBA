Sub AddNewColumns()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("KPI Briefing Date  2024 w CCs")

    Dim i As Integer
    For i = 1 To 7
        ws.Columns("AA:AA").Insert Shift:=xlToRight
    Next i

    ws.Range("AA1").Value = "CCID"
    ws.Range("AB1").Value = "FIN 4"
    ws.Range("AC1").Value = "Fin 2"
    ws.Range("AD1").Value = "Fin 3"
    ws.Range("AE1").Value = "Fin 5"
    ws.Range("AF1").Value = "Fin 7"
    ws.Range("AG1").Value = "GEO"

    ws.Range("Z:Z").Copy
    ws.Range("AA:AG").PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

                ws.Range("AA2").Formula = "=IF(IFERROR(IFERROR(VALUE(LEFT(IF(LEFT(RIGHT(Y2,10),1)=""("",RIGHT(Y2,9),RIGHT(Y2,11)),LEN(IF(LEFT(RIGHT(Y2,10),1)=""("",RIGHT(Y2,9),RIGHT(Y2,11)))-1)),LEFT(IF(LEFT(RIGHT(Y2,10),1)=""("",RIGHT(Y2,9),RIGHT(Y2,11)),LEN(IF(LEFT(RIGHT(Y2,10),1)=""("",RIGHT(Y2,9),RIGHT(Y2,11)))-1)),Z2)=0,""Unmapped"",IFERROR(IFERROR(VALUE(LEFT(IF(LEFT(RIGHT(Y2,10),1)=""("",RIGHT(Y2,9),RIGHT(Y2,11)),LEN(IF(LEFT(RIGHT(Y2,10),1)=""("",RIGHT(Y2,9),RIGHT(Y2,11)))-1)),LEFT(IF(LEFT(RIGHT(Y2,10),1)=""("",RIGHT(Y2,9),RIGHT(Y2,11)),LEN(IF(LEFT(RIGHT(Y2,10),1)=""("",RIGHT(Y2,9),RIGHT(Y2,11)))-1)),Z2))"
ws.Range("AA2:AA" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row).FillDown

ws.Range("AA:AA").Value = ws.Range("AA:AA").Value
End Sub

Sub ApplyTextToColumns()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("KPI Briefing Date  2024 w CCs")

     With ws
        .Columns("B").TextToColumns Destination:=.Range("B1"), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
        .Columns("G").TextToColumns Destination:=.Range("G1"), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
        .Columns("X").TextToColumns Destination:=.Range("X1"), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True

        .Columns("J").Replace What:=" (*)", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        .Columns("I").Replace What:="??-", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

        .Range("AB2").Formula = "=IFNA(VLOOKUP(AA2, [CC_MAP.xlsx]Sheet1!$A:$G,7,0),IF(LEFT(J2,1)=""A"",""Atos"",""Eviden""))"
        .Range("AC2").Formula = "=IFNA(VLOOKUP(AA2, [CC_MAP.xlsx]Sheet1!$A:$E,5,0),XLOOKUP(ad2,[CC_MAP.xlsx]Sheet1!$F:$F,[CC_MAP.xlsx]Sheet1!$E:$E))"
        .Range("AD2").Formula = "=IFNA(VLOOKUP(AA2, [CC_MAP.xlsx]Sheet1!$A:$F,6,0),I2)"
        .Range("AE2").Formula = "=IFNA(VLOOKUP(AA2, [CC_MAP.xlsx]Sheet1!$A:$H,8,0),XLOOKUP(ad2,[CC_MAP.xlsx]Sheet1!$F:$F,[CC_MAP.xlsx]Sheet1!$H:$H))"
        .Range("AF2").Formula = "=IFNA(VLOOKUP(AA2, [CC_MAP.xlsx]Sheet1!$A:$J,10,0),XLOOKUP(ad2,[CC_MAP.xlsx]Sheet1!$F:$F,[CC_MAP.xlsx]Sheet1!$J:$J))"
        .Range("AG2").Formula = "=IFNA(XLOOKUP(aa2,[CC_MAP.xlsx]Sheet1!$A:$A,[CC_MAP.xlsx]Sheet1!$D:$D),XLOOKUP(ad2,[CC_MAP.xlsx]Sheet1!$F:$F,[CC_MAP.xlsx]Sheet1!$D:$D))"

    ws.Range("AB2:AB" & ws.Cells(ws.Rows.Count, "AA").End(xlUp).Row).FillDown
    ws.Range("AC2:AC" & ws.Cells(ws.Rows.Count, "AA").End(xlUp).Row).FillDown
    ws.Range("AD2:AD" & ws.Cells(ws.Rows.Count, "AA").End(xlUp).Row).FillDown
    ws.Range("AE2:AE" & ws.Cells(ws.Rows.Count, "AA").End(xlUp).Row).FillDown
    ws.Range("AF2:AF" & ws.Cells(ws.Rows.Count, "AA").End(xlUp).Row).FillDown
    ws.Range("AG2:AG" & ws.Cells(ws.Rows.Count, "AA").End(xlUp).Row).FillDown

   ws.Range("AB2:AG" & ws.Cells(ws.Rows.Count, "AA").End(xlUp).Row).FillDown
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

    Set DSheet = Worksheets("KPI Briefing Date  2024 w CCs")

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

