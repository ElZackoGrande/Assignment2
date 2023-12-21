Sub Assignment2()
'
' Assignment2 Macro
'

'
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "Ticker"
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "Yearly Change"
    Range("K1").Select
    ActiveCell.FormulaR1C1 = "Percent Change"
    Range("L1").Select
    ActiveCell.FormulaR1C1 = "Total Stock Volume"
    Columns("A:A").Select
    Selection.Copy
    Columns("I:I").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveSheet.Range("$I$1:$L$753001").RemoveDuplicates Columns:=1, Header:= _
        xlNo
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "Ticker"
    Range("J2").Select
    ActiveCell.FormulaR1C1 = _
        "=MAXIFS(C6,C1,RC9,C2,MAX(C2))-MINIFS(C3,C1,RC9,C2,MIN(C2))"
    Range("B2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWindow.ScrollRow = 1048547
    ActiveWindow.ScrollRow = 1026037
    ActiveWindow.ScrollRow = 982895
    ActiveWindow.ScrollRow = 493324
    ActiveWindow.ScrollRow = 433299
    ActiveWindow.ScrollRow = 380778
    ActiveWindow.ScrollRow = 210085
    ActiveWindow.ScrollRow = 196955
    ActiveWindow.ScrollRow = 174446
    ActiveWindow.ScrollRow = 78782
    ActiveWindow.ScrollRow = 73155
    ActiveWindow.ScrollRow = 67528
    ActiveWindow.ScrollRow = 54397
    ActiveWindow.ScrollRow = 52522
    ActiveWindow.ScrollRow = 33764
    ActiveWindow.ScrollRow = 31888
    ActiveWindow.ScrollRow = 30013
    ActiveWindow.ScrollRow = 13131
    ActiveWindow.ScrollRow = 11255
    ActiveWindow.ScrollRow = 9379
    ActiveWindow.ScrollRow = 5628
    ActiveWindow.ScrollRow = 3752
    ActiveWindow.ScrollRow = 1876
    ActiveWindow.ScrollRow = 1
    Range("J2").Select
    ActiveCell.FormulaR1C1 = _
        "=MAXIFS(C6,C1,RC9,C2,MAX(C2))-MINIFS(C3,C1,RC9,C2,MIN(C2))"
    Range("J2").Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 5287936
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.Copy
    Range("K2").Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = _
        "=MAXIFS(C6,C1,RC9,C2,MAX(C2))/MINIFS(C3,C1,RC9,C2,MIN(C2))-1"
    Range("K2").Select
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.0%"
    Selection.NumberFormat = "0.00%"
    Range("L2").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(C1,RC9,C7)"
    Range("J2:L2").Select
    Selection.AutoFill Destination:=Range("J2:L3001")
    Range("J2:L3001").Select
    Columns("I:L").Select
    Columns("I:L").EntireColumn.AutoFit
    Range("O2").Select
    ActiveCell.FormulaR1C1 = "Greatest % Increase"
    Range("O3").Select
    ActiveCell.FormulaR1C1 = "Greatest % Decrease"
    Range("O4").Select
    ActiveCell.FormulaR1C1 = "Greatest Total Volume"
    Range("O5").Select
    Columns("O:O").EntireColumn.AutoFit
    Range("P1").Select
    ActiveCell.FormulaR1C1 = "Ticker"
    Range("Q1").Select
    ActiveCell.FormulaR1C1 = "Value"
    Range("Q2").Select
    ActiveCell.FormulaR1C1 = "=MAX(C11)"
    Range("Q3").Select
    ActiveCell.FormulaR1C1 = "=MIN(C[-6])"
    Range("Q4").Select
    ActiveCell.FormulaR1C1 = "=MAX(C[-5])"
    Range("Q2:Q3").Select
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.0%"
    Selection.NumberFormat = "0.00%"
    Range("P2").Select
    ActiveCell.FormulaR1C1 = "=INDEX(C9,MATCH(R[1]C17,C11,0))"
    Range("P2").Select
    Selection.AutoFill Destination:=Range("P2:P4"), Type:=xlFillDefault
    Range("P2:P4").Select
    Range("P4").Select
    ActiveCell.FormulaR1C1 = "=INDEX(C9,MATCH(R[1]C17,C12,0))"
    Range("Q2").Select
    ActiveCell.FormulaR1C1 = "=MAX(C11)"
    Range("P2").Select
    ActiveCell.FormulaR1C1 = "=INDEX(C9,MATCH(R[1]C17,C11,0))"
    Range("P2").Select
    ActiveCell.FormulaR1C1 = "=INDEX(C9,MATCH(RC17,C11,0))"
    Range("P3").Select
    ActiveCell.FormulaR1C1 = "=INDEX(C9,MATCH(RC17,C11,0))"
    Range("P4").Select
    ActiveCell.FormulaR1C1 = "=INDEX(C9,MATCH(RC17,C12,0))"
    Columns("I:Q").Select
    Selection.Copy
    Sheets("2019").Select
    Columns("I:I").Select
    ActiveSheet.Paste
    Sheets("2020").Select
    Columns("I:I").Select
    ActiveSheet.Paste
    Sheets("2019").Select
    Range("B2").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWindow.ScrollRow = 743801
    ActiveWindow.ScrollRow = 670774
    ActiveWindow.ScrollRow = 435462
    ActiveWindow.ScrollRow = 416529
    ActiveWindow.ScrollRow = 403005
    ActiveWindow.ScrollRow = 319159
    ActiveWindow.ScrollRow = 306987
    ActiveWindow.ScrollRow = 243426
    ActiveWindow.ScrollRow = 239369
    ActiveWindow.ScrollRow = 236664
    ActiveWindow.ScrollRow = 204208
    ActiveWindow.ScrollRow = 201503
    ActiveWindow.ScrollRow = 197446
    ActiveWindow.ScrollRow = 163637
    ActiveWindow.ScrollRow = 160932
    ActiveWindow.ScrollRow = 159580
    ActiveWindow.ScrollRow = 133885
    ActiveWindow.ScrollRow = 128475
    ActiveWindow.ScrollRow = 114952
    ActiveWindow.ScrollRow = 25695
    ActiveWindow.ScrollRow = 10819
    ActiveWindow.ScrollRow = 5410
    ActiveWindow.ScrollRow = 1
    Sheets("2020").Select
    Range("B2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWindow.ScrollRow = 1048547
    ActiveWindow.ScrollRow = 1042919
    ActiveWindow.ScrollRow = 1005404
    ActiveWindow.ScrollRow = 934126
    ActiveWindow.ScrollRow = 761556
    ActiveWindow.ScrollRow = 315127
    ActiveWindow.ScrollRow = 309500
    ActiveWindow.ScrollRow = 300121
    ActiveWindow.ScrollRow = 290742
    ActiveWindow.ScrollRow = 286991
    ActiveWindow.ScrollRow = 285115
    ActiveWindow.ScrollRow = 281363
    ActiveWindow.ScrollRow = 277612
    ActiveWindow.ScrollRow = 273860
    ActiveWindow.ScrollRow = 245724
    ActiveWindow.ScrollRow = 240097
    ActiveWindow.ScrollRow = 195079
    ActiveWindow.ScrollRow = 189452
    ActiveWindow.ScrollRow = 181949
    ActiveWindow.ScrollRow = 138806
    ActiveWindow.ScrollRow = 131303
    ActiveWindow.ScrollRow = 120049
    ActiveWindow.ScrollRow = 112546
    ActiveWindow.ScrollRow = 65652
    ActiveWindow.ScrollRow = 54397
    ActiveWindow.ScrollRow = 18758
    ActiveWindow.ScrollRow = 13131
    ActiveWindow.ScrollRow = 7504
    ActiveWindow.ScrollRow = 1
    Range("P17").Select
    Application.CutCopyMode = False
    Calculate
End Sub
