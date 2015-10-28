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


Sub columnSearch()
'
' columnSearch Macro
'

'
    'Range("TestComments[[#Headers],[Discussion Points]]").Select
    'Range(Selection, Selection.End(xlDown)).Select
    'Range(Selection, Selection.End(xlDown)).Select
    'Range(Selection, Selection.End(xlDown)).Select
    'ActiveWindow.SmallScroll Down:=6
    
    Set thetable = ActiveWorkbook.ActiveSheet.ListObjects("TestComments")
    thetable.ListColumns("Discussion Points").DataBodyRange.Select
    
    
    Selection.Find(What:="*", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
        
    Selection.FindNext(After:=ActiveCell).Activate
    Selection.FindNext(After:=ActiveCell).Activate
    ActiveWindow.SmallScroll Down:=-21
    Selection.FindNext(After:=ActiveCell).Activate
    Selection.FindNext(After:=ActiveCell).Activate
End Sub


