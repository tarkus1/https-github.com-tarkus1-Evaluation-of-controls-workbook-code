Attribute VB_Name = "Module5"
Sub hidecomments()
Attribute hidecomments.VB_ProcData.VB_Invoke_Func = " \n14"
'
' hidecomments Macro
'

'
    Application.DisplayCommentIndicator = xlCommentAndIndicator
    Application.DisplayCommentIndicator = xlCommentIndicatorOnly
End Sub
Sub hideacomment()
Attribute hideacomment.VB_ProcData.VB_Invoke_Func = " \n14"
'
' hideacomment Macro
'

'
    Range("D11").Select
    ActiveCell.Comment.Visible = True
End Sub
