Sub AddNewColumns()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("KPI Time to Recruit 2024 w CCs ")

    Dim i As Integer
    For i = 1 To 7
        ws.Columns("AE:AK").Insert Shift:=xlToRight
    Next i

    ws.Range("AE1").Value = "CCID"
    ws.Range("AF1").Value = "FIN 4"
    ws.Range("AG1").Value = "Fin 2"
    ws.Range("AH1").Value = "Fin 3"
    ws.Range("AI1").Value = "Fin 5"
    ws.Range("AJ1").Value = "Fin 7"
    ws.Range("AK1").Value = "GEO"

    ws.Range("Z:Z").Copy
    ws.Range("AE:AK").PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

        ws.Range("AE2").Formula = "=IF(IFERROR(IFERROR(VALUE(LEFT(IF(LEFT(RIGHT(AC2,10),1)=""("",RIGHT(AC2,9),RIGHT(AC2,11)),LEN(IF(LEFT(RIGHT(AC2,10),1)=""("",RIGHT(AC2,9),RIGHT(AC2,11)))-1)),LEFT(IF(LEFT(RIGHT(AC2,10),1)=""("",RIGHT(AC2,9),RIGHT(AC2,11)),LEN(IF(LEFT(RIGHT(AC2,10),1)=""("",RIGHT(AC2,9),RIGHT(AC2,11)))-1)),AD2)=0,""Unmapped"",IFERROR(IFERROR(VALUE(LEFT(IF(LEFT(RIGHT(AC2,10),1)=""("",RIGHT(AC2,9),RIGHT(AC2,11)),LEN(IF(LEFT(RIGHT(AC2,10),1)=""("",RIGHT(AC2,9),RIGHT(AC2,11)))-1)),LEFT(IF(LEFT(RIGHT(AC2,10),1)=""("",RIGHT(AC2,9),RIGHT(AC2,11)),LEN(IF(LEFT(RIGHT(AC2,10),1)=""("",RIGHT(AC2,9),RIGHT(AC2,11)))-1)),AD2))"
ws.Range("AE2:AE" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row).FillDown

ws.Range("AE:AE").Value = ws.Range("AE:AE").Value
End Sub

Sub ApplyTextToColumns()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("KPI Time to Recruit 2024 w CCs ")

     With ws
        .Columns("U").TextToColumns Destination:=.Range("U1"), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
        .Columns("K").TextToColumns Destination:=.Range("K1"), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
        .Columns("N").TextToColumns Destination:=.Range("N1"), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
        .Columns("AB").TextToColumns Destination:=.Range("AB1"), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True


        .Columns("H").Replace What:=" (*)", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        .Columns("E").Replace What:="??-", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

        .Range("AF2").Formula = "=IFNA(VLOOKUP(AE2, [CC_MAP.xlsx]Sheet2!$A:$G,7,0),IF(LEFT(H2,1)=""A"",""Atos"",""Eviden""))"
        .Range("AG2").Formula = "=IFNA(VLOOKUP(AE2, [CC_MAP.xlsx]Sheet1!$A:$E,5,0),XLOOKUP(ah2,[CC_MAP.xlsx]Sheet1!$F:$F,[CC_MAP.xlsx]Sheet1!$E:$E))"
        .Range("AH2").Formula = "=IFNA(VLOOKUP(AE2, [CC_MAP.xlsx]Sheet1!$A:$F,6,0),E2)"
        .Range("AI2").Formula = "=IFNA(VLOOKUP(AE2, [CC_MAP.xlsx]Sheet1!$A:$H,8,0),XLOOKUP(ah2,[CC_MAP.xlsx]Sheet1!$F:$F,[CC_MAP.xlsx]Sheet1!$H:$H))"
        .Range("AJ2").Formula = "=IFNA(VLOOKUP(AE2, [CC_MAP.xlsx]Sheet1!$A:$J,10,0),XLOOKUP(ah2,[CC_MAP.xlsx]Sheet1!$F:$F,[CC_MAP.xlsx]Sheet1!$J:$J))"
        .Range("AK2").Formula = "=IFNA(XLOOKUP(AE2,[CC_MAP.xlsx]Sheet1!$A:$A,[CC_MAP.xlsx]Sheet1!$D:$D),XLOOKUP(AH2,[CC_MAP.xlsx]Sheet1!$F:$F,[CC_MAP.xlsx]Sheet1!$D:$D))"

    ws.Range("AF2:AF" & ws.Cells(ws.Rows.Count, "AE").End(xlUp).Row).FillDown
    ws.Range("AG2:AG" & ws.Cells(ws.Rows.Count, "AE").End(xlUp).Row).FillDown
    ws.Range("AH2:AH" & ws.Cells(ws.Rows.Count, "AE").End(xlUp).Row).FillDown
    ws.Range("AI2:AI" & ws.Cells(ws.Rows.Count, "AE").End(xlUp).Row).FillDown
    ws.Range("AJ2:AJ" & ws.Cells(ws.Rows.Count, "AE").End(xlUp).Row).FillDown
    ws.Range("AK2:AK" & ws.Cells(ws.Rows.Count, "AE").End(xlUp).Row).FillDown

   ws.Range("AF2:AK" & ws.Cells(ws.Rows.Count, "AE").End(xlUp).Row).FillDown
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

    Set DSheet = Worksheets("KPI Time to Recruit 2024 w CCs ")

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

