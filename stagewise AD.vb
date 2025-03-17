Sub AddNewColumnsWithFormulaAndFormatPainter()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Stage Wise Application UPDATED ")
    
    Dim i As Integer
    For i = 1 To 10
        ws.Columns("AP:AP").Insert Shift:=xlToRight
    Next i
    
    ws.Range("AP1").Value = "CCID"
    ws.Range("AQ1").Value = "Fin 4"
    ws.Range("AR1").Value = "Fin 2"
    ws.Range("AS1").Value = "Fin 3"
    ws.Range("AT1").Value = "Fin 5 - Business Line"
    ws.Range("AU1").Value = "Fin 7 - Sub-branch"
    ws.Range("AV1").Value = "GEO"
    ws.Range("AW1").Value = "rec"
    ws.Range("AX1").Value = "supp"
    ws.Range("AY1").Value = "supp team"
    
    ws.Range("AO:AO").Copy
    ws.Range("AP:AY").PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
    
    ws.Range("AP2").Formula = "=IF(IFERROR(IFERROR(VALUE(LEFT(IF(LEFT(RIGHT(AL2,10),1)=""("",RIGHT(AL2,9),RIGHT(AL2,11)),LEN(IF(LEFT(RIGHT(AL2,10),1)=""("",RIGHT(AL2,9),RIGHT(AL2,11)))-1)),LEFT(IF(LEFT(RIGHT(AL2,10),1)=""("",RIGHT(AL2,9),RIGHT(AL2,11)),LEN(IF(LEFT(RIGHT(AL2,10),1)=""("",RIGHT(AL2,9),RIGHT(AL2,11)))-1)),AM2)=0,""Unmapped"",IFERROR(IFERROR(VALUE(LEFT(IF(LEFT(RIGHT(AL2,10),1)=""("",RIGHT(AL2,9),RIGHT(AL2,11)),LEN(IF(LEFT(RIGHT(AL2,10),1)=""("",RIGHT(AL2,9),RIGHT(AL2,11)))-1)),LEFT(IF(LEFT(RIGHT(AL2,10),1)=""("",RIGHT(AL2,9),RIGHT(AL2,11)),LEN(IF(LEFT(RIGHT(AL2,10),1)=""("",RIGHT(AL2,9),RIGHT(AL2,11)))-1)),AM2))"
    
    ws.Range("AP2:AP" & ws.Cells(ws.Rows.Count, "AL").End(xlUp).Row).FillDown
    
    ws.Range("AP:AP").Value = ws.Range("AP:AP").Value
End Sub

Sub ApplyAllOperations()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Stage Wise Application UPDATED ")
    
    With ws
        .Columns("M").Replace What:="??-", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        
        .Columns("N").Replace What:=" (*)", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        .Columns("O").Replace What:=" (*)", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        
        .Columns("AG").TextToColumns Destination:=.Range("AG1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), _
        TrailingMinusNumbers:=True
        
        .Columns("AN").TextToColumns Destination:=.Range("AN1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), _
        TrailingMinusNumbers:=True
        
        .Columns("AO").TextToColumns Destination:=.Range("AO1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), _
        TrailingMinusNumbers:=True

        .Columns("AZ").TextToColumns Destination:=.Range("AZ1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), _
        TrailingMinusNumbers:=True
        
        .Range("AQ2").Formula = "=IFNA(XLOOKUP(ap2, [CC_MAP.xlsx]Sheet1!$A:$A, [CC_MAP.xlsx]Sheet1!$G:$G), IF(LEFT(o2, 1) = ""A"", ""Atos"", ""Eviden""))"
        .Range("AR2").Formula = "=IFNA(VLOOKUP(AP2,[CC_MAP.xlsx]Sheet1!$A:$E,5,0),XLOOKUP(AS2,[CC_MAP.xlsx]Sheet1!$F:$F,[CC_MAP.xlsx]Sheet1!$E:$E))"
        .Range("AS2").Formula = "=IFNA(VLOOKUP(AP2,[CC_MAP.xlsx]Sheet1!$A:$F,6,0),m2)"
        .Range("AT2").Formula = "=IFNA(VLOOKUP(AP2,[CC_MAP.xlsx]Sheet1!$A:$H,8,0),XLOOKUP(AS2,[CC_MAP.xlsx]Sheet1!$F:$F,[CC_MAP.xlsx]Sheet1!$H:$H))"
        .Range("AU2").Formula = "=IFNA(VLOOKUP(AP2,[CC_MAP.xlsx]Sheet1!$A:$J,10,0),XLOOKUP(AS2,[CC_MAP.xlsx]Sheet1!$F:$F,[CC_MAP.xlsx]Sheet1!$J:$J))"
        .Range("AV2").Formula = "=IFNA(VLOOKUP(AP2,[CC_MAP.xlsx]Sheet1!$A:$M,4,0),XLOOKUP(ar2,[CC_MAP.xlsx]Sheet1!$e:$e,[CC_MAP.xlsx]Sheet1!$d:$d))"
        .Range("AW2").Formula = "=IFNA(VLOOKUP(AN2,'[MasterList_with CC.xlsx]MasterList'!$D:$AJ,19,FALSE),IFNA(VLOOKUP(AN2,'[MasterList_with CC.xlsx]MasterList'!$C:$AJ,20,FALSE),IFNA(VLOOKUP(AN2,'[MasterList_with CC.xlsx]MasterList'!$B:$AJ,21,FALSE),VLOOKUP(AN2,'[MasterList_with CC.xlsx]MasterList'!$D:$AJ,19,FALSE))))"
        .Range("AX2").Formula = "=IFNA(VLOOKUP(AO2,'[MasterList_with CC.xlsx]MasterList'!$D:$AJ,19,FALSE),IFNA(VLOOKUP(AO2,'[MasterList_with CC.xlsx]MasterList'!$C:$AJ,20,FALSE),IFNA(VLOOKUP(AO2,'[MasterList_with CC.xlsx]MasterList'!$B:$AJ,21,FALSE),VLOOKUP(AO2,'[MasterList_with CC.xlsx]MasterList'!$D:$AJ,19,FALSE))))"
        .Range("AY2").Formula = "=IFNA(VLOOKUP(AO2,'[MasterList_with CC.xlsx]MasterList'!$D:$AJ,5,FALSE),IFNA(VLOOKUP(AO2,'[MasterList_with CC.xlsx]MasterList'!$C:$AJ,6,FALSE),IFNA(VLOOKUP(AO2,'[MasterList_with CC.xlsx]MasterList'!$B:$AJ,7,FALSE),VLOOKUP(AO2,'[MasterList_with CC.xlsx]MasterList'!$D:$AJ,5,FALSE))))"
        

ws.Range("AQ2:AQ" & ws.Cells(ws.Rows.Count, "AP").End(xlUp).Row).FillDown
ws.Range("AR2:AR" & ws.Cells(ws.Rows.Count, "AP").End(xlUp).Row).FillDown
ws.Range("AS2:AS" & ws.Cells(ws.Rows.Count, "AP").End(xlUp).Row).FillDown
ws.Range("AT2:AT" & ws.Cells(ws.Rows.Count, "AP").End(xlUp).Row).FillDown
ws.Range("AU2:AU" & ws.Cells(ws.Rows.Count, "AP").End(xlUp).Row).FillDown
ws.Range("AV2:AV" & ws.Cells(ws.Rows.Count, "AP").End(xlUp).Row).FillDown
ws.Range("AW2:AW" & ws.Cells(ws.Rows.Count, "AP").End(xlUp).Row).FillDown
ws.Range("AX2:AX" & ws.Cells(ws.Rows.Count, "AP").End(xlUp).Row).FillDown
ws.Range("AY2:AY" & ws.Cells(ws.Rows.Count, "AP").End(xlUp).Row).FillDown

ws.Range("AQ2:AY" & ws.Cells(ws.Rows.Count, "AP").End(xlUp).Row).Value = ws.Range("AQ2:AY" & ws.Cells(ws.Rows.Count, "AP").End(xlUp).Row).Value
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

    Set DSheet = Worksheets("Stage Wise Application UPDATED ")

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

