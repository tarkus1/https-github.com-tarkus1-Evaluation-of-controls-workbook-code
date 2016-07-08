Attribute VB_Name = "CommentSizeSnippets"

Sub sizeComment()
Attribute sizeComment.VB_ProcData.VB_Invoke_Func = " \n14"
'
' sizeComment Macro
'

'
    Range("D14").Comment.Text Text:= _
        "Accurate well status determined and maintained in PA system and Petrinex (producing, shut in, suspended or abandoned)."
    Selection.ShapeRange.ScaleWidth 1.15, msoFalse, msoScaleFromTopLeft
    Selection.ShapeRange.ScaleHeight 1.53, msoFalse, msoScaleFromTopLeft
    Range("D14").Select
End Sub
Sub tryAgain()
Attribute tryAgain.VB_ProcData.VB_Invoke_Func = " \n14"
'
' tryAgain Macro
'

'
    
    For Each chago In ActiveSheet.Shapes
        Debug.Print chago.Name
        ' Application.WorksheetFunction.Search("Comment", chago.Name)
        
        
    Next chago
    
End Sub
