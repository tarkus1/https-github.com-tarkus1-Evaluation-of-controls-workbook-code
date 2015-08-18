Attribute VB_Name = "Module3"
Sub copyNCE()
Attribute copyNCE.VB_ProcData.VB_Invoke_Func = " \n14"
'
' copyNCE Macro
'

'
    Application.CutCopyMode = False
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$T$1:$Y$29"), , xlYes).Name = _
        "Table27"
    Range("U2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("BPT3 - Gas MeasurementTest").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveWindow.SmallScroll Down:=18
    Selection.End(xlDown).Select
    Range("A39").Select
    Sheets("Client Controls").Select
    Range("Q2").Select
    Application.CutCopyMode = False
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$P$1:$S$9"), , xlYes).Name = _
        "Table28"
    Range("Table28").Select
    Selection.Copy
    Sheets("BPT3 - Gas MeasurementTest").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("E38").Select
    ActiveWindow.SmallScroll Down:=-12
End Sub

Sub text2number()
Attribute text2number.VB_ProcData.VB_Invoke_Func = " \n14"
'
' text2number Macro
'

'
    Range("A11").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.NumberFormat = "0"
    Range("F9").Select
    
End Sub
Sub PageBreakReset()
Attribute PageBreakReset.VB_ProcData.VB_Invoke_Func = " \n14"
'
' PageBreakReset Macro
'

'
    ActiveSheet.PageSetup.PrintArea = "$C$1:$I$48"
 
End Sub
Sub Macro9()
Attribute Macro9.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro9 Macro
'

'
    Range("U2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("X22").Select
    ActiveWindow.SmallScroll Down:=-18
    Range("BP_NCEs[[Theme]:[NCE Risk]]").Select
       Range("GasMeas27[[#Headers],[Theme]]").Select
    ActiveWindow.SmallScroll Down:=-9
    Range("A12").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
   Range("GasMeas27[Theme]").Select
End Sub
Sub Macro11()
Attribute Macro11.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro11 Macro
'

'
    Range("GasMeas27[[#Headers],[Theme]]").Select
    Selection.AutoFilter
    ActiveWorkbook.Worksheets("BPT3 - Gas MeasurementTest").ListObjects("GasMeas27" _
        ).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("BPT3 - Gas MeasurementTest").ListObjects("GasMeas27" _
        ).Sort.SortFields.Add Key:=Range("GasMeas27[[#All],[Theme]]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("BPT3 - Gas MeasurementTest").ListObjects( _
        "GasMeas27").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("BPT3 - Gas MeasurementTest").ListObjects("GasMeas27" _
        ).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("BPT3 - Gas MeasurementTest").ListObjects("GasMeas27" _
        ).Sort.SortFields.Add Key:=Range("GasMeas27[[#All],[NCE]]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("BPT3 - Gas MeasurementTest").ListObjects( _
        "GasMeas27").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

Sub tablerange()
    ActiveWorkbook.Worksheets("BPT3 - Gas MeasurementTest").ListObjects(1).ListColumns("NCE").DataBodyRange.Select

End Sub


Sub Macro13()
Attribute Macro13.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro13 Macro
'

'
    Range("GasMeas27[[#Headers],[NCE Component]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlTop
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
End Sub
