Attribute VB_Name = "CodeSnippets"
Public addCount As Integer

Sub doit()
    InsertTableColumn
    SetColumnWidth
    GetIDs
    TableHeaderFill
End Sub
Sub InsertTableColumn()
Attribute InsertTableColumn.VB_ProcData.VB_Invoke_Func = " \n14"
'
' InsertTableColumn Macro
'
    addCount = 1
    Do While counter < addCount
        Selection.ListObject.ListColumns.Add Position:=29
        '    Range("GasExFac[[#Headers],[ABBT0052144]]").Select
        counter = counter + 1
    Loop
End Sub
Sub SetColumnWidth()
Attribute SetColumnWidth.VB_ProcData.VB_Invoke_Func = " \n14"
'
' SetColumnWidth Macro
'
    Columns("I:BV").Select
    Range("I4").Activate
    Selection.ColumnWidth = 17
End Sub
Sub GetIDs()
Attribute GetIDs.VB_ProcData.VB_Invoke_Func = " \n14"
'
' GetIDs Macro
'
Dim fromSheet As Worksheet
    
    Set fromSheet = ActiveWorkbook.ActiveSheet
    Application.Goto Reference:="FacIds"
    Selection.Copy
    
    fromSheet.Activate
    
    Range("i10").Select
      
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub

Sub TableHeaderFill()
Attribute TableHeaderFill.VB_ProcData.VB_Invoke_Func = " \n14"
'
' TableHeaderFill Macro
'

'
    Range("I9").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("I2:BV9").Select
    Range("I9").Activate
    Selection.FillRight
End Sub
Sub TableStartSelect()
Attribute TableStartSelect.VB_ProcData.VB_Invoke_Func = " \n14"
'
' TableStartSelect Macro
'

'
    Range("GasNFac[[#Headers]]").Select
End Sub


Sub TableFormFill()
'
' TableFormFill Macro
'
 Dim wrksht As Worksheet
 Dim oListObj As ListObject
 
' Set wrksht = ActiveWorkbook.Worksheets("BP1 - Gas Exist Fac Des & Inst")
  Set wrksht = ActiveWorkbook.Worksheets("Test Sheet")

 wrksht.Activate
 
 Set oListObj = wrksht.ListObjects(1)
 
 'MsgBox oListObj.Name
 
 MsgBox oListObj.ListColumns.Count
 
 oListObj.ListColumns(4).DataBodyRange.Select
 oListObj.ListColumns(4).Delete
 
 ' MsgBox oListObj.ListColumns.
 'MsgBox oListObj.ListColumns.Item.Name
 
 
 
 'oListObj.ListColumns(9).Range.Select
  ' oListObj.ListColumns(9).DataBodyRange(2).Select

 


' this is not reliable!
'    Range("I11").Select
'    Range(Selection, Selection.End(xlDown)).Select
'    Range(Selection, Selection.End(xlToRight)).Select
'    Range(Selection, Selection.End(xlToRight)).Select
'    Selection.FillRight
End Sub
