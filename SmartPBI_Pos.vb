Sub AddNewColumns()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Gold_Requisition_Dump_Juan Copy")
    
    Dim i As Integer
    For i = 1 To 2
        ws.Columns("AE:AE").Insert Shift:=xlToRight
    Next i

    ws.Range("AE1").Value = "CCID"
    ws.Range("AF1").Value = "Fin 4"

    ws.Range("AD:AD").Copy
    ws.Range("AE:AF").PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

    ws.Range("AE2").Formula = "=IF(IFERROR(IFERROR(VALUE(LEFT(IF(LEFT(RIGHT(AC2,10),1)=""("",RIGHT(AC2,9),RIGHT(AC2,11)),LEN(IF(LEFT(RIGHT(AC2,10),1)=""("",RIGHT(AC2,9),RIGHT(AC2,11)))-1)),LEFT(IF(LEFT(RIGHT(AC2,10),1)=""("",RIGHT(AC2,9),RIGHT(AC2,11)),LEN(IF(LEFT(RIGHT(AC2,10),1)=""("",RIGHT(AC2,9),RIGHT(AC2,11)))-1)),AD2)=0,""Unmapped"",IFERROR(IFERROR(VALUE(LEFT(IF(LEFT(RIGHT(AC2,10),1)=""("",RIGHT(AC2,9),RIGHT(AC2,11)),LEN(IF(LEFT(RIGHT(AC2,10),1)=""("",RIGHT(AC2,9),RIGHT(AC2,11)))-1)),LEFT(IF(LEFT(RIGHT(AC2,10),1)=""("",RIGHT(AC2,9),RIGHT(AC2,11)),LEN(IF(LEFT(RIGHT(AC2,10),1)=""("",RIGHT(AC2,9),RIGHT(AC2,11)))-1)),AD2))"
ws.Range("AE2:AE" & ws.Cells(ws.Rows.Count, "AB").End(xlUp).Row).FillDown

ws.Range("AE:AE").Value = ws.Range("AE:AE").Value
End Sub

Sub ApplyTextToColumns()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Gold_Requisition_Dump_Juan Copy")
    
    With ws
        .Columns("B").TextToColumns Destination:=.Range("B1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), _
        TrailingMinusNumbers:=True
        
        .Columns("C").TextToColumns Destination:=.Range("C1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), _
        TrailingMinusNumbers:=True
        
        .Columns("G").TextToColumns Destination:=.Range("G1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), _
        TrailingMinusNumbers:=True
        
        .Columns("X").TextToColumns Destination:=.Range("X1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), _
        TrailingMinusNumbers:=True
        
        .Columns("Y").TextToColumns Destination:=.Range("Y1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), _
        TrailingMinusNumbers:=True

         .Columns("Z").TextToColumns Destination:=.Range("Z1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), _
        TrailingMinusNumbers:=True

        .Columns("U").Replace What:=" (*)", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        .Columns("V").Replace What:=" (*)", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        
        .Columns("T").Replace What:="??-", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        .Columns("M").Replace What:="* - ", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

        .Range("AF2").Formula = "=IFNA(IFS(VLOOKUP(AE2, 'C:\Users\a880862\Downloads\[CC_MAP.xlsx]Sheet1'!$A:$G, 7, FALSE) = ""Global Tech"", ""Atos"", VLOOKUP(AE2, 'C:\Users\a880862\Downloads\[CC_MAP.xlsx]Sheet1'!$A:$G, 7, FALSE) = ""Global Tech Support"", ""Atos"", TRUE, VLOOKUP(AE2, 'C:\Users\a880862\Downloads\[CC_MAP.xlsx]Sheet1'!$A:$G, 7, FALSE)), IF(LEFT(U2, 1) = ""A"", ""A"", ""E""))"
        ws.Range("AF2:AF" & ws.Cells(ws.Rows.Count, "AB").End(xlUp).Row).FillDown
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

    Set DSheet = Worksheets("Gold_Requisition_Dump_Juan Copy")

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
        
    With .PivotFields("Job Req ID")
            .Orientation = xlDataField
            .Function = xlCount
            .Name = "Count of Job Req ID"
    End With
End With
End Sub