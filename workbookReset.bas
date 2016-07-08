Attribute VB_Name = "workbookReset"
Dim theSheet As Worksheet
'
' This module should be deleted from client copies to avoid accidentally
' destroying the results



' step through all sheets
Sub stepThroughReset()

Dim wsSheets As Worksheets, wsActive As Worksheet, wsSheet As Worksheet

' Set wsSheet = ActiveWorkbook.Worksheets
Application.Calculation = xlManual

Rebuild = True


For Each wsSheet In ActiveWorkbook.Worksheets

    If Left(wsSheet.Name, 2) = "BP" Then
        
            Set theSheet = wsSheet
            MsgBox theSheet.Name
            
            removeColumns
            removeHeader
            clearConcEvid
            
    End If
Next wsSheet

Application.Calculation = xlAutomatic
Rebuild = False

End Sub


' remove all but the 1st and 2nd twp last facility columns
Sub removeColumns()

    'Dim wrksht As Worksheet
    Dim oListObj As ListObject
    Dim columnCount As Variant
 
    'Set wrksht = ActiveWorkbook.Worksheets("Test Sheet")

    theSheet.Activate
    'MsgBox theSheet.Name
    
    Set oListObj = theSheet.ListObjects(1)
 
    Do Until oListObj.ListColumns.Count < 13
        columnCount = oListObj.ListColumns.Count
        Debug.Print columnCount
        oListObj.ListColumns(11).Delete
    Loop
 

End Sub

' remove extra header info

Sub removeHeader()

' based on L9 being in the first column beyond the reduced table

    theSheet.Activate
    'MsgBox theSheet.Name

 Range("k9").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    
    Selection.Clear

End Sub

'clear the conclusion and evidence

Sub clearConcEvid()

    Dim wrksht As Worksheet
    Dim oListObj As ListObject
    Dim reasonCol As Range
 
    Set theSheet = ActiveWorkbook.ActiveSheet

 theSheet.Activate
 
 Set oListObj = theSheet.ListObjects(1)
 

 
 oListObj.ListColumns("Conclusion").DataBodyRange.ClearContents
 

 oListObj.ListColumns("Evidence").DataBodyRange.ClearContents
 
 oListObj.ListColumns("Control Performer Role").DataBodyRange.ClearContents
  
 oListObj.ListColumns("Reason for Conclusion").DataBodyRange.ClearContents
  
 ' need a way to preserve the N/A formula
 ' Set reasonCol = oListObj.ListColumns("Reason for Conclusion").DataBodyRange

    ' For Each cell In reasonCol
        ' If Left(cell.Formula, 4) <> "=IF(" Then
        ' cell.ClearContents
        ' End If
    'Next cell
End Sub

