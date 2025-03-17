Sub AddNewColumns()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Strategic Hires")
    
    Dim i As Integer
    For i = 1 To 14
        ws.Columns("AY:AY").Insert Shift:=xlToRight
    Next i
    
    ws.Range("AY1").Value = "CCID"
    ws.Range("AZ1").Value = "Fin 4"
    ws.Range("BA1").Value = "Fin 2"
    ws.Range("BB1").Value = "Fin 3"
    ws.Range("BC1").Value = "Fin 5 - Business Line"
    ws.Range("BD1").Value = "GEO"
    ws.Range("BE1").Value = "rec"
    ws.Range("BF1").Value = "supp"
    ws.Range("BG1").Value = "supp team"
    ws.Range("BH1").Value = "rec group"
    ws.Range("BI1").Value = "supp group"
    ws.Range("BJ1").Value = "gcm split"
    ws.Range("BK1").Value = "Last Status Changed Date Months"
    ws.Range("BL1").Value = "Fin 7"

    
    ws.Range("AT:AT").Copy
    ws.Range("AY:BL").PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

    ws.Range("AY2").Formula = "=IF(IFERROR(IFERROR(VALUE(LEFT(IF(LEFT(RIGHT(AJ2,10),1)=""("",RIGHT(AJ2,9),RIGHT(AJ2,11)),LEN(IF(LEFT(RIGHT(AJ2,10),1)=""("",RIGHT(AJ2,9),RIGHT(AJ2,11)))-1)),LEFT(IF(LEFT(RIGHT(AJ2,10),1)=""("",RIGHT(AJ2,9),RIGHT(AJ2,11)),LEN(IF(LEFT(RIGHT(AJ2,10),1)=""("",RIGHT(AJ2,9),RIGHT(AJ2,11)))-1)),AK2)=0,""Unmapped"",IFERROR(IFERROR(VALUE(LEFT(IF(LEFT(RIGHT(AJ2,10),1)=""("",RIGHT(AJ2,9),RIGHT(AJ2,11)),LEN(IF(LEFT(RIGHT(AJ2,10),1)=""("",RIGHT(AJ2,9),RIGHT(AJ2,11)))-1)),LEFT(IF(LEFT(RIGHT(AJ2,10),1)=""("",RIGHT(AJ2,9),RIGHT(AJ2,11)),LEN(IF(LEFT(RIGHT(AJ2,10),1)=""("",RIGHT(AJ2,9),RIGHT(AJ2,11)))-1)),AK2))"
ws.Range("AY2:AY" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row).FillDown

ws.Range("AY:AY").Value = ws.Range("AY:AY").Value
End Sub

Sub ApplyTextToColumns()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Strategic Hires")
    
    With ws
        .Columns("B").TextToColumns Destination:=.Range("B1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), _
        TrailingMinusNumbers:=True
        
        .Columns("C").TextToColumns Destination:=.Range("C1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), _
        TrailingMinusNumbers:=True
        
        .Columns("V").TextToColumns Destination:=.Range("V1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), _
        TrailingMinusNumbers:=True
        
        .Columns("AD").TextToColumns Destination:=.Range("AD1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), _
        TrailingMinusNumbers:=True
        
        .Columns("AE").TextToColumns Destination:=.Range("AE1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), _
        TrailingMinusNumbers:=True

        .Columns("AX").TextToColumns Destination:=.Range("AX1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), _
        TrailingMinusNumbers:=True

        .Columns("AB").Replace What:=" (*)", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        
        .Columns("Z").Replace What:="??-", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

ws.Range("AZ2").Formula = "=IFNA(IFS(VLOOKUP(AU2,[CC_MAP.xlsx]Sheet1!$A:$G,7,0)=""Global Tech"", ""Atos"", VLOOKUP(AU2,[CC_MAP.xlsx]Sheet1!$A:$G,7,0)=""Global Tech Support"", ""Atos""), IF(LEFT(AB2,1)=""A"",""Atos"",""Eviden""))"
ws.Range("BA2").Formula = "=IFNA(VLOOKUP(AY2,[CC_MAP.xlsx]Sheet1!$A:$E,5,0),XLOOKUP(AX2,[CC_MAP.xlsx]Sheet1!$F:$F,[CC_MAP.xlsx]Sheet1!$E:$E))"
ws.Range("BB2").Formula = "=IFNA(VLOOKUP(AY2,[CC_MAP.xlsx]Sheet1!$A:$F,6,0),Z2)"
ws.Range("BC2").Formula = "=IFNA(VLOOKUP(AY2,[CC_MAP.xlsx]Sheet1!$A:$H,8,0),XLOOKUP(AX2,[CC_MAP.xlsx]Sheet1!$F:$F,[CC_MAP.xlsx]Sheet1!$H:$H))"
ws.Range("BD2").Formula = "=IFNA(VLOOKUP(AY2,[CC_MAP.xlsx]Sheet1!$A:$M,13,0),XLOOKUP(AX2,[CC_MAP.xlsx]Sheet1!$F:$F,[CC_MAP.xlsx]Sheet1!$M:$M))"
ws.Range("BE2").Formula = "=IFNA(VLOOKUP(AD2,'[MasterList_with CC.xlsx]MasterList'!$D:$AJ,19,FALSE),IFNA(VLOOKUP(AD2,'[MasterList_with CC.xlsx]MasterList'!$C:$AJ,20,FALSE),IFNA(VLOOKUP(AD2,'[MasterList_with CC.xlsx]MasterList'!$B:$AJ,21,FALSE),VLOOKUP(AD2,'[MasterList_with CC.xlsx]MasterList'!$D:$AJ,19,FALSE))))"
ws.Range("BF2").Formula = "=IFNA(VLOOKUP(AE2,'[MasterList_with CC.xlsx]MasterList'!$D:$AJ,19,FALSE),IFNA(VLOOKUP(AE2,'[MasterList_with CC.xlsx]MasterList'!$C:$AJ,20,FALSE),IFNA(VLOOKUP(AE2,'[MasterList_with CC.xlsx]MasterList'!$B:$AJ,21,FALSE),VLOOKUP(AE2,'[MasterList_with CC.xlsx]MasterList'!$D:$AJ,19,FALSE))))"
ws.Range("BG2").Formula = "=IFNA(VLOOKUP(AE2,'[MasterList_with CC.xlsx]MasterList'!$D:$AJ,5,FALSE),IFNA(VLOOKUP(AE2,'[MasterList_with CC.xlsx]MasterList'!$C:$AJ,6,FALSE),IFNA(VLOOKUP(AE2,'[MasterList_with CC.xlsx]MasterList'!$B:$AJ,7,FALSE),VLOOKUP(AE2,'[MasterList_with CC.xlsx]MasterList'!$D:$AJ,5,FALSE))))"
ws.Range("BH2").Formula = "=IFNA(VLOOKUP(AD2,'[MasterList_with CC.xlsx]MasterList'!$D:$AJ,17,FALSE),IFNA(VLOOKUP(AD2,'[MasterList_with CC.xlsx]MasterList'!$C:$AJ,18,FALSE),IFNA(VLOOKUP(AD2,'[MasterList_with CC.xlsx]MasterList'!$B:$AJ,19,FALSE),VLOOKUP(AD2,'[MasterList_with CC.xlsx]MasterList'!$D:$AJ,17,FALSE))))"
ws.Range("BI2").Formula = "=IFNA(VLOOKUP(AE2,'[MasterList_with CC.xlsx]MasterList'!$D:$AJ,17,FALSE),IFNA(VLOOKUP(AE2,'[MasterList_with CC.xlsx]MasterList'!$C:$AJ,18,FALSE),IFNA(VLOOKUP(AE2,'[MasterList_with CC.xlsx]MasterList'!$B:$AJ,19,FALSE),VLOOKUP(AE2,'[MasterList_with CC.xlsx]MasterList'!$D:$AJ,17,FALSE))))"
ws.Range("BJ2").Formula = "=IFERROR(IFS(V2=""N/A"",""No GCM Level"",VALUE(V2) <=3,""GCM 0-3"",VALUE(V2) <=6,""GCM 4-6"",VALUE(V2) >=7,""GCM 7+""),""No GCM Level"")"
ws.Range("BK2").Formula = "=TEXT(AQ2,""mmm"")"
ws.Range("BL2").Formula = "=IFNA(VLOOKUP(AY2,[CC_MAP.xlsx]Sheet1!$A:$J,10,0),XLOOKUP(AX2,[CC_MAP.xlsx]Sheet1!$F:$F,[CC_MAP.xlsx]Sheet1!$J:$J))"

ws.Range("AZ2:AZ" & ws.Cells(ws.Rows.Count, "AU").End(xlUp).Row).FillDown
ws.Range("BA2:BA" & ws.Cells(ws.Rows.Count, "AU").End(xlUp).Row).FillDown
ws.Range("BB2:BB" & ws.Cells(ws.Rows.Count, "AU").End(xlUp).Row).FillDown
ws.Range("BC2:BC" & ws.Cells(ws.Rows.Count, "AU").End(xlUp).Row).FillDown
ws.Range("BD2:BD" & ws.Cells(ws.Rows.Count, "AU").End(xlUp).Row).FillDown
ws.Range("BE2:BE" & ws.Cells(ws.Rows.Count, "AU").End(xlUp).Row).FillDown
ws.Range("BF2:BF" & ws.Cells(ws.Rows.Count, "AU").End(xlUp).Row).FillDown
ws.Range("BG2:BG" & ws.Cells(ws.Rows.Count, "AU").End(xlUp).Row).FillDown
ws.Range("BH2:BH" & ws.Cells(ws.Rows.Count, "AU").End(xlUp).Row).FillDown
ws.Range("BI2:BI" & ws.Cells(ws.Rows.Count, "AU").End(xlUp).Row).FillDown
ws.Range("BJ2:BJ" & ws.Cells(ws.Rows.Count, "AU").End(xlUp).Row).FillDown
ws.Range("BK2:BK" & ws.Cells(ws.Rows.Count, "AU").End(xlUp).Row).FillDown
ws.Range("BL2:BL" & ws.Cells(ws.Rows.Count, "AU").End(xlUp).Row).FillDown

ws.Range("AZ2:BL" & ws.Cells(ws.Rows.Count, "AU").End(xlUp).Row).Value = ws.Range("AZ2:BL" & ws.Cells(ws.Rows.Count, "AU").End(xlUp).Row).Value

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

    Set DSheet = Worksheets("Strategic Hires")

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

