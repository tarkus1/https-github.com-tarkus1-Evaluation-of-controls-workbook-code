Attribute VB_Name = "Module6"
Sub FormatNA()
Attribute FormatNA.VB_ProcData.VB_Invoke_Func = " \n14"
'
' FormatNA Macro
'

'
    With Selection.Font
        .Name = "Arial"
        .FontStyle = "Regular"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .Color = 255
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 14013951
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
End Sub
Sub ClearFormat()
Attribute ClearFormat.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ClearFormat Macro
'

'
    Selection.ClearFormats
End Sub
