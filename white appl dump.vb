Sub ApplyAllValues()
     Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Gold_Applications_Dump_2023 GOL")
    
    Dim i As Integer
    For i = 1 To 5
        ws.Columns("AX:AX").Insert Shift:=xlToRight
    Next i
    
    ws.Range("AX1").Value = "CCID"
    ws.Range("AY1").Value = "Fin 4"
    ws.Range("AZ1").Value = "Final Rec SYS ID"
    ws.Range("BA1").Value = "Final Rec Type"
    ws.Range("BB1").Value = "In Calculation"
    
    ws.Range("AU:AU").Copy
    ws.Range("AX:BB").PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

    ws.Range("AY2").Formula = "=IF(IFERROR(IFERROR(VALUE(LEFT(IF(LEFT(RIGHT(Z2,10),1)=""("",RIGHT(Z2,9),RIGHT(Z2,11)),LEN(IF(LEFT(RIGHT(Z2,10),1)=""("",RIGHT(Z2,9),RIGHT(Z2,11)))-1)),LEFT(IF(LEFT(RIGHT(Z2,10),1)=""("",RIGHT(Z2,9),RIGHT(Z2,11)),LEN(IF(LEFT(RIGHT(Z2,10),1)=""("",RIGHT(Z2,9),RIGHT(Z2,11)))-1)),Y2)=0,""Unmapped"",IFERROR(IFERROR(VALUE(LEFT(IF(LEFT(RIGHT(Z2,10),1)=""("",RIGHT(Z2,9),RIGHT(Z2,11)),LEN(IF(LEFT(RIGHT(Z2,10),1)=""("",RIGHT(Z2,9),RIGHT(Z2,11)))-1)),LEFT(IF(LEFT(RIGHT(Z2,10),1)=""("",RIGHT(Z2,9),RIGHT(Z2,11)),LEN(IF(LEFT(RIGHT(Z2,10),1)=""("",RIGHT(Z2,9),RIGHT(Z2,11)))-1)),Y2))"
ws.Range("AY2:AY" & ws.Cells(ws.Rows.Count, "AU").End(xlUp).Row).FillDown

ws.Range("AY2").Formula = "=IFNA(IFS(VLOOKUP(AX2,[CC_MAP.xlsx]Sheet1!$A:$G,7,0)=""Global Tech"", ""Atos"", VLOOKUP(AX2,[CC_MAP.xlsx]Sheet1!$A:$G,7,0)=""Global Tech Support"", ""Atos""), IF(LEFT(AG2, 1) = ""A"", ""Atos"", ""Not Found""))"
ws.Range("AZ2").Formula = "=IF(V2="""",Q2,V2)"
ws.Range("BA2").Formula = "=IFNA(VLOOKUP(AZ2,'[MasterList_withCC.xlsx]MasterList'!$D:$AJ,19,FALSE), IFNA(VLOOKUP(AZ2,'[MasterList_withCC.xlsx]MasterList'!$C:$AJ,20,FALSE), IFNA(VLOOKUP(AZ2,'[MasterList_withCC.xlsx]MasterList'!$B:$AJ,21,FALSE), ""N/A"")))"
ws.Range("BB2").Formula = "=IF(BA2=""N/A"",""NO"",""YES"")"
ws.Range("AY2:AY" & ws.Cells(ws.Rows.Count, "AX").End(xlUp).Row).FillDown
ws.Range("AZ2:AZ" & ws.Cells(ws.Rows.Count, "AX").End(xlUp).Row).FillDown
ws.Range("BA2:BA" & ws.Cells(ws.Rows.Count, "AX").End(xlUp).Row).FillDown
ws.Range("BB2:BB" & ws.Cells(ws.Rows.Count, "AX").End(xlUp).Row).FillDown


End Sub
