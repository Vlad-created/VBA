Sub AddNewColumnsWithFormulaAndFormatPainter()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Offers_Test Copy")  
    
    Dim i As Integer
    For i = 1 To 6
        ws.Columns("AL:AL").Insert Shift:=xlToRight
    Next i
    
    ws.Range("AL1").Value = "CCID"
    ws.Range("AM1").Value = "Fin 4"
    ws.Range("AN1").Value = "Fin 2"
    ws.Range("AO1").Value = "Fin 3"
    ws.Range("AP1").Value = "Fin 5 - Business Line"
    ws.Range("AQ1").Value = "GEO"
    ws.Range("AR1").Value = "TTO"
    ws.Range("AS1").Value = "Salary Gap"
    

    
    With ws
        .Columns("F").TextToColumns Destination:=.Range("F1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), _
        TrailingMinusNumbers:=True
        
        .Columns("AA").TextToColumns Destination:=.Range("AA1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), _
        TrailingMinusNumbers:=True
        
        .Columns("AD").TextToColumns Destination:=.Range("AD1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), _
        TrailingMinusNumbers:=True
        
        .Columns("AG").TextToColumns Destination:=.Range("AG1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), _
        TrailingMinusNumbers:=True
        
        .Columns("Q").Replace What:=" (*)", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        .Columns("R").Replace What:=" (*)", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        
        .Columns("E").Replace What:="??-", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    
    ws.Range("AL2").Formula = "=IF(IFERROR(IFERROR(VALUE(LEFT(IF(LEFT(RIGHT(AK2,10),1)=""("",RIGHT(AK2,9),RIGHT(AK2,11)),LEN(IF(LEFT(RIGHT(AK2,10),1)=""("",RIGHT(AK2,9),RIGHT(AK2,11)))-1)),LEFT(IF(LEFT(RIGHT(AK2,10),1)=""("",RIGHT(AK2,9),RIGHT(AK2,11)),LEN(IF(LEFT(RIGHT(AK2,10),1)=""("",RIGHT(AK2,9),RIGHT(AK2,11)))-1)),AJ2)=0,""Unmapped"",IFERROR(IFERROR(VALUE(LEFT(IF(LEFT(RIGHT(AK2,10),1)=""("",RIGHT(AK2,9),RIGHT(AK2,11)),LEN(IF(LEFT(RIGHT(AK2,10),1)=""("",RIGHT(AK2,9),RIGHT(AK2,11)))-1)),LEFT(IF(LEFT(RIGHT(AK2,10),1)=""("",RIGHT(AK2,9),RIGHT(AK2,11)),LEN(IF(LEFT(RIGHT(AK2,10),1)=""("",RIGHT(AK2,9),RIGHT(AK2,11)))-1)),AJ2))"
    
    ws.Range("AM2").Formula = "=IFNA(VLOOKUP(AL2,'C:\Users\a880862\Downloads\[CC_MAP.xlsx]Sheet1'!$A:$G,7,0),IF(LEFT(Q2,1)=""A"",""Atos"",""Eviden""))"    
    ws.Range("AN2").Formula = "=IFNA(VLOOKUP(AL2,[CC_MAP.xlsx]Sheet1!$A:$E,5,0),XLOOKUP(AO2,[CC_MAP.xlsx]Sheet1!$F:$F,[CC_MAP.xlsx]Sheet1!$E:$E))"
    ws.Range("AO2").Formula = "=IFNA(VLOOKUP(AL2,[CC_MAP.xlsx]Sheet1!$A:$F,6,0),E2)"
    ws.Range("AP2").Formula = "=IFNA(VLOOKUP(AL2,[CC_MAP.xlsx]Sheet1!$A:$H,8,0),XLOOKUP(AO2,[CC_MAP.xlsx]Sheet1!$F:$F,[CC_MAP.xlsx]Sheet1!$H:$H))"
    ws.Range("AQ2").Formula = "=IFNA(VLOOKUP(AL2,[CC_MAP.xlsx]Sheet1!$A:$M,4,0),XLOOKUP(an2,[CC_MAP.xlsx]Sheet1!$e:$e,[CC_MAP.xlsx]Sheet1!$d:$d))"


    ws.Range("AK:AK").Copy
    ws.Range("AL:AQ").PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
    End With
End Sub

Sub ApplyFormulasAndFormat()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Offers_Test Copy")
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "K").End(xlUp).Row

    ws.Range("AR2:AR" & lastRow).Formula = "=IF(K2="""",""No date"",K2-J2)"

    ws.Range("AS2:AS" & lastRow).Formula = "=IF(Y2=0,""no date"",(X2-Y2)/Y2)"

    ws.Range("AR2:AR" & lastRow).NumberFormat = "0"

    ws.Range("AS2:AS" & lastRow).NumberFormat = "0%"

    ws.Range("AL2:AQ" & lastRow).FillDown
  
    ws.Range("AL2:AS" & lastRow).Value = ws.Range("AL2:AS" & lastRow).Value

    ws.Range("AK:AK").Copy
    ws.Range("AL:AS").PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
End Sub

