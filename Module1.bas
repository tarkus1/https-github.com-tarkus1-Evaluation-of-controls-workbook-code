Attribute VB_Name = "Module1"

Sub copyFormat()
Attribute copyFormat.VB_ProcData.VB_Invoke_Func = " \n14"
'
' copyFormat Macro
'

'
    Selection.Copy
    Range("C12:C24").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
End Sub
