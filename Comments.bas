Attribute VB_Name = "Comments"
Sub Comments()
Attribute Comments.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Comments Macro
'

'
    Range("GasExFac[[#Headers],[NCE Component]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.End(xlDown).Select
    Range("C28").Select
    ActiveCell.FormulaR1C1 = "Comments:"
    .Select


    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection.Font
        .Name = "Arial"
        .Size = 8
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
End Sub
