Attribute VB_Name = "Module7"
Sub shade()
Attribute shade.VB_ProcData.VB_Invoke_Func = " \n14"
'
' shade Macro
'

'
    Range("I22").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub
