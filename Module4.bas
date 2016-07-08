Attribute VB_Name = "Module4"
Sub textformat()
Attribute textformat.VB_ProcData.VB_Invoke_Func = " \n14"
'
' textformat Macro
'

'
    Range("D11").Select
    Selection.ShapeRange.SetShapesDefaultProperties
    Selection.ShapeRange.SetShapesDefaultProperties
    Range("D11").Comment.Text Text:= _
        "Are you operating delivery point measurement or are you relying on the downstream receiver for delivery point measurement? How do you know the measurement device is installed and used correctly?"
    Range("D11").Select
End Sub
