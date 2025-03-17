Sub AddNewColumns()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Gold_Applications_Dump_Juan Cop")
    
    Dim i As Integer
    For i = 1 To 2
        ws.Columns("AI:AI").Insert Shift:=xlToRight
    Next i

    ws.Range("AI1").Value = "CCID"
    ws.Range("AJ1").Value = "Fin 4"

    ws.Range("AD:AD").Copy
    ws.Range("AI:AJ").PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

    ws.Range("AI2").Formula = "=IF(IFERROR(IFERROR(VALUE(LEFT(IF(LEFT(RIGHT(AG2,10),1)=""("",RIGHT(AG2,9),RIGHT(AG2,11)),LEN(IF(LEFT(RIGHT(AG2,10),1)=""("",RIGHT(AG2,9),RIGHT(AG2,11)))-1)),LEFT(IF(LEFT(RIGHT(AG2,10),1)=""("",RIGHT(AG2,9),RIGHT(AG2,11)),LEN(IF(LEFT(RIGHT(AG2,10),1)=""("",RIGHT(AG2,9),RIGHT(AG2,11)))-1)),AH2)=0,""Unmapped"",IFERROR(IFERROR(VALUE(LEFT(IF(LEFT(RIGHT(AG2,10),1)=""("",RIGHT(AG2,9),RIGHT(AG2,11)),LEN(IF(LEFT(RIGHT(AG2,10),1)=""("",RIGHT(AG2,9),RIGHT(AG2,11)))-1)),LEFT(IF(LEFT(RIGHT(AG2,10),1)=""("",RIGHT(AG2,9),RIGHT(AG2,11)),LEN(IF(LEFT(RIGHT(AG2,10),1)=""("",RIGHT(AG2,9),RIGHT(AG2,11)))-1)),AH2))"
ws.Range("AI2:AI" & ws.Cells(ws.Rows.Count, "AF").End(xlUp).Row).FillDown

ws.Range("AI:AI").Value = ws.Range("AI:AI").Value
End Sub

Sub ApplyTextToColumns()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Gold_Applications_Dump_Juan Cop")
    
    With ws
        .Columns("B").TextToColumns Destination:=.Range("B1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), _
        TrailingMinusNumbers:=True
        
        .Columns("C").TextToColumns Destination:=.Range("C1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), _
        TrailingMinusNumbers:=True
        
        .Columns("J").TextToColumns Destination:=.Range("J1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), _
        TrailingMinusNumbers:=True
        
        .Columns("R").TextToColumns Destination:=.Range("R1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), _
        TrailingMinusNumbers:=True
        
        .Columns("S").TextToColumns Destination:=.Range("S1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), _
        TrailingMinusNumbers:=True

        .Columns("X").TextToColumns Destination:=.Range("X1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), _
        TrailingMinusNumbers:=True
        
        .Columns("Q").TextToColumns Destination:=.Range("Q1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), _
        TrailingMinusNumbers:=True

        .Columns("W").Replace What:=" (*)", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        .Columns("V").Replace What:=" (*)", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        
        .Columns("T").Replace What:="??-", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        .Columns("K").Replace What:="* - ", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False


        .Range("AJ2").Formula = "=IFNA(IFS(VLOOKUP(AI2, '[CC_MAP.xlsx]Sheet1'!$A:$G, 7, FALSE) = ""Global Tech"", ""Atos"", VLOOKUP(AI2, '[CC_MAP.xlsx]Sheet1'!$A:$G, 7, FALSE) = ""Global Tech Support"", ""Atos"", TRUE, VLOOKUP(AI2, '[CC_MAP.xlsx]Sheet1'!$A:$G, 7, FALSE)), IF(LEFT(W2, 1) = ""A"", ""A"", ""E""))"
        ws.Range("AJ2:AJ" & ws.Cells(ws.Rows.Count, "AB").End(xlUp).Row).FillDown
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

    Set DSheet = Worksheets("Gold_Applications_Dump_Juan Cop")

    LastRow = DSheet.Cells(DSheet.Rows.Count, 1).End(xlUp).Row
    LastCol = DSheet.Cells(1, DSheet.Columns.Count).End(xlToLeft).Column
    Set PRange = DSheet.Cells(1, 1).Resize(LastRow, LastCol)

    Set PCache = ActiveWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=PRange)

    Set PSheet = Worksheets.Add

    Set PTable = PCache.CreatePivotTable( _
        TableDestination:=PSheet.Cells(1, 1), _
        TableName:="PivotTable1")

    With PTable
    End With

    With PTable
        .PivotFields("Fin 4").Orientation = xlRowField
        
    With .PivotFields("Application ID")
            .Orientation = xlDataField
            .Function = xlCount
            .Name = "Count of Application ID"
    End With
End With
End Sub