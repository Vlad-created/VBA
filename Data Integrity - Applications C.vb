Sub AddNewColumns()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Data Integrity - Applications C")  
    
    Dim i As Integer
    For i = 1 To 7
        ws.Columns("AB:AB").Insert Shift:=xlToRight
    Next i
    
    ws.Range("AB1").Value = "CCID"
    ws.Range("AC1").Value = "Fin 4"
    ws.Range("AD1").Value = "Fin 2"
    ws.Range("AE1").Value = "Fin 3"
    ws.Range("AF1").Value = "Fin 5 - Business Line"
    ws.Range("AG1").Value = "Fin 7 - Sub-branch"
    ws.Range("AH1").Value = "GEO"
    
    ws.Range("AA:AA").Copy
    ws.Range("AB:AH").PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

    ws.Range("AB2").Formula = "=IF(IFERROR(IFERROR(VALUE(LEFT(IF(LEFT(RIGHT(X2,10),1)=""("",RIGHT(X2,9),RIGHT(X2,11)),LEN(IF(LEFT(RIGHT(X2,10),1)=""("",RIGHT(X2,9),RIGHT(X2,11)))-1)),LEFT(IF(LEFT(RIGHT(X2,10),1)=""("",RIGHT(X2,9),RIGHT(X2,11)),LEN(IF(LEFT(RIGHT(X2,10),1)=""("",RIGHT(X2,9),RIGHT(X2,11)))-1)),Y2)=0,""Unmapped"",IFERROR(IFERROR(VALUE(LEFT(IF(LEFT(RIGHT(X2,10),1)=""("",RIGHT(X2,9),RIGHT(X2,11)),LEN(IF(LEFT(RIGHT(X2,10),1)=""("",RIGHT(X2,9),RIGHT(X2,11)))-1)),LEFT(IF(LEFT(RIGHT(X2,10),1)=""("",RIGHT(X2,9),RIGHT(X2,11)),LEN(IF(LEFT(RIGHT(X2,10),1)=""("",RIGHT(X2,9),RIGHT(X2,11)))-1)),Y2))"

ws.Range("AB2:AH" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row).FillDown

ws.Range("AB:AB").Value = ws.Range("AB:AB").Value

    With ws
        .Columns("Z").TextToColumns Destination:=.Range("Z1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), _
        TrailingMinusNumbers:=True
        .Columns("AA").TextToColumns Destination:=.Range("AA1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), _
        TrailingMinusNumbers:=True
        .Columns("AI").TextToColumns Destination:=.Range("AI1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), _
        TrailingMinusNumbers:=True
        .Columns("AJ").TextToColumns Destination:=.Range("AJ1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), _
        TrailingMinusNumbers:=True

        .Columns("O").Replace What:=" (*)", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        .Columns("P").Replace What:=" (*)", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        
        .Columns("Q").Replace What:="??-", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        End With
End Sub

Sub AddVLOOKUPAndCopyAsValues()

    Dim LastRow As Long

    With ThisWorkbook.Sheets("Data Integrity - Applications C") 
        .Range("AC2").Formula = "=IFNA(XLOOKUP(AB2,[CC_MAP.xlsx]Sheet1!$A:$A,[CC_MAP.xlsx]Sheet1!$G:$G), IF(LEFT(P2, 1) = ""A"", ""Atos"", ""Eviden""))"
        .Range("AD2").Formula = "=IFNA(VLOOKUP(AB2,'C:\Users\a880862\Downloads\[CC_MAP.xlsx]Sheet1'!$A:$E,5,0),XLOOKUP(AE2,'C:\Users\a880862\Downloads\[CC_MAP.xlsx]Sheet1'!$F:$F,'C:\Users\a880862\Downloads\[CC_MAP.xlsx]Sheet1'!$E:$E))"
        .Range("AE2").Formula = "=IFNA(VLOOKUP(AB2,'C:\Users\a880862\Downloads\[CC_MAP.xlsx]Sheet1'!$A:$F,6,0),Q2)"
        .Range("AF2").Formula = "=IFNA(VLOOKUP(AB2,'C:\Users\a880862\Downloads\[CC_MAP.xlsx]Sheet1'!$A:$H,8,0),XLOOKUP(AE2,'C:\Users\a880862\Downloads\[CC_MAP.xlsx]Sheet1'!$F:$F,'C:\Users\a880862\Downloads\[CC_MAP.xlsx]Sheet1'!$H:$H))"
        .Range("AG2").Formula = "=IFNA(VLOOKUP(AB2,'C:\Users\a880862\Downloads\[CC_MAP.xlsx]Sheet1'!$A:$J,10,0),XLOOKUP(AE2,'C:\Users\a880862\Downloads\[CC_MAP.xlsx]Sheet1'!$F:$F,'C:\Users\a880862\Downloads\[CC_MAP.xlsx]Sheet1'!$J:$J))"
        .Range("AH2").Formula = "=IFNA(XLOOKUP(AB2,'[CC_MAP.xlsx]Sheet1'!$A:$A,'[CC_MAP.xlsx]Sheet1'!$D:$D),XLOOKUP(ad2,'[CC_MAP.xlsx]Sheet1'!$e:$e,'[CC_MAP.xlsx]Sheet1'!$D:$D))"

        LastRow = .Cells(.Rows.Count, "AB").End(xlUp).Row

        .Range("AC2:AC" & LastRow).FillDown
        .Range("AD2:AD" & LastRow).FillDown
        .Range("AE2:AE" & LastRow).FillDown
        .Range("AF2:AF" & LastRow).FillDown
        .Range("AG2:AG" & LastRow).FillDown
        .Range("AH2:AH" & LastRow).FillDown

        .Range("AB2:AH" & LastRow).Copy
        .Range("AB2:AH" & LastRow).PasteSpecial xlPasteValues
        Application.CutCopyMode = False
    End With

End Sub

