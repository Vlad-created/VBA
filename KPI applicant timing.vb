Sub AddNewColumns()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("KPI New Applicant Timing 2024 w")

    Dim i As Integer
    For i = 1 To 7
        ws.Columns("AG:AM").Insert Shift:=xlToRight
    Next i

    ws.Range("AG1").Value = "CCID"
    ws.Range("AH1").Value = "FIN 4"
    ws.Range("AI1").Value = "Fin 2"
    ws.Range("AJ1").Value = "Fin 3"
    ws.Range("AK1").Value = "Fin 5"
    ws.Range("AL1").Value = "Fin 7"
    ws.Range("AM1").Value = "GEO"

    ws.Range("Z:Z").Copy
    ws.Range("AG:AM").PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

                        ws.Range("AG2").Formula = "=IF(IFERROR(IFERROR(VALUE(LEFT(IF(LEFT(RIGHT(AE2,10),1)=""("",RIGHT(AE2,9),RIGHT(AE2,11)),LEN(IF(LEFT(RIGHT(AE2,10),1)=""("",RIGHT(AE2,9),RIGHT(AE2,11)))-1)),LEFT(IF(LEFT(RIGHT(AE2,10),1)=""("",RIGHT(AE2,9),RIGHT(AE2,11)),LEN(IF(LEFT(RIGHT(AE2,10),1)=""("",RIGHT(AE2,9),RIGHT(AE2,11)))-1)),AF2)=0,""Unmapped"",IFERROR(IFERROR(VALUE(LEFT(IF(LEFT(RIGHT(AE2,10),1)=""("",RIGHT(AE2,9),RIGHT(AE2,11)),LEN(IF(LEFT(RIGHT(AE2,10),1)=""("",RIGHT(AE2,9),RIGHT(AE2,11)))-1)),LEFT(IF(LEFT(RIGHT(AE2,10),1)=""("",RIGHT(AE2,9),RIGHT(AE2,11)),LEN(IF(LEFT(RIGHT(AE2,10),1)=""("",RIGHT(AE2,9),RIGHT(AE2,11)))-1)),AF2))"
ws.Range("AG2:AG" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row).FillDown

ws.Range("AG:AG").Value = ws.Range("AG:AG").Value
End Sub

Sub ApplyTextToColumns()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("KPI New Applicant Timing 2024 w")

     With ws
        .Columns("B").TextToColumns Destination:=.Range("B1"), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
        .Columns("F").TextToColumns Destination:=.Range("F1"), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
        .Columns("N").TextToColumns Destination:=.Range("N1"), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
        .Columns("Q").TextToColumns Destination:=.Range("Q1"), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
        .Columns("AD").TextToColumns Destination:=.Range("AD1"), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True


        .Columns("I").Replace What:=" (*)", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        .Columns("H").Replace What:="??-", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

        .Range("AH2").Formula = "=IFNA(VLOOKUP(AG2, [CC_MAP.xlsx]Sheet1!$A:$G,7,0),IF(LEFT(I2,1)=""A"",""Atos"",""Eviden""))"
        .Range("AI2").Formula = "=IFNA(VLOOKUP(AG2, [CC_MAP.xlsx]Sheet1!$A:$E,5,0),XLOOKUP(aj2,[CC_MAP.xlsx]Sheet1!$F:$F,[CC_MAP.xlsx]Sheet1!$E:$E))"
        .Range("AJ2").Formula = "=IFNA(VLOOKUP(AG2, [CC_MAP.xlsx]Sheet1!$A:$F,6,0),H2)"
        .Range("AK2").Formula = "=IFNA(VLOOKUP(AG2, [CC_MAP.xlsx]Sheet1!$A:$H,8,0),XLOOKUP(aj2,[CC_MAP.xlsx]Sheet1!$F:$F,[CC_MAP.xlsx]Sheet1!$H:$H))"
        .Range("AL2").Formula = "=IFNA(VLOOKUP(AG2, [CC_MAP.xlsx]Sheet1!$A:$J,10,0),XLOOKUP(aj2,[CC_MAP.xlsx]Sheet1!$F:$F,[CC_MAP.xlsx]Sheet1!$J:$J))"
        .Range("AM2").Formula = "=IFNA(XLOOKUP(ag2,[CC_MAP.xlsx]Sheet1!$A:$A,[CC_MAP.xlsx]Sheet1!$D:$D),XLOOKUP(aj2,[CC_MAP.xlsx]Sheet1!$F:$F,[CC_MAP.xlsx]Sheet1!$D:$D))"

    ws.Range("AH2:AH" & ws.Cells(ws.Rows.Count, "AG").End(xlUp).Row).FillDown
    ws.Range("AI2:AI" & ws.Cells(ws.Rows.Count, "AG").End(xlUp).Row).FillDown
    ws.Range("AJ2:AJ" & ws.Cells(ws.Rows.Count, "AG").End(xlUp).Row).FillDown
    ws.Range("AK2:AK" & ws.Cells(ws.Rows.Count, "AG").End(xlUp).Row).FillDown
    ws.Range("AL2:AL" & ws.Cells(ws.Rows.Count, "AG").End(xlUp).Row).FillDown
    ws.Range("AM2:AM" & ws.Cells(ws.Rows.Count, "AG").End(xlUp).Row).FillDown

   ws.Range("AH2:AM" & ws.Cells(ws.Rows.Count, "AG").End(xlUp).Row).FillDown
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

    Set DSheet = Worksheets("KPI New Applicant Timing 2024 w")

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

