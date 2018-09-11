Attribute VB_Name = "workbookBuild"
Dim theSheet As Worksheet, facCount As Variant


' step through all BP sheets
Sub stepThroughBuild()

Dim wsSheets As Worksheets, wsActive As Worksheet, wsSheet As Worksheet

' set the rebuild variable
Rebuild = True

columnNumber

For Each wsSheet In ActiveWorkbook.Worksheets

    If Left(wsSheet.Name, 2) = "BP" Then
        
            Set theSheet = wsSheet
            MsgBox theSheet.Name
            insertColumns
            insertHeaderInfo
            fillTableForumulas
            ' NAConclusion
            
        
    End If
Next wsSheet

NAConclusion
concFormat
rowHeights

' done, reset the rebuild varable
Rebuild = False

End Sub
' figure out how many columns to add from the facility list

Sub columnNumber()
    Dim wrksht As Worksheet
    Set wrksht = ActiveWorkbook.Worksheets("Facility List")
    wrksht.Activate
    
    Range(Range("B18"), Range("B18").End(xlToRight)).Name = "FacIDs"
    
  
    
    wrksht.Range("FacIDs").Select
    
   facCount = wrksht.Range("FacIDs").Count
    Debug.Print "facility count "; facCount

End Sub

' paste in the new facility IDs
Sub insertColumns()
    Dim oListObj As ListObject, headerRange As Range
    
    ' Set theSheet = ActiveSheet
    
    Application.GoTo Reference:="FacIDs"
    Selection.Copy
    theSheet.Activate
    
    Set oListObj = theSheet.ListObjects(1)
    
    ' MsgBox oListObj.ListColumns(9).Name
    
    oListObj.HeaderRowRange.Offset(0, 9).Activate
    
    ' headerRange.Offset(0, 8).Activate
    ActiveCell.PasteSpecial (xlPasteValues)
    

End Sub


' fill the facility information rows above
Sub insertHeaderInfo()
'
' insertHeaderInfo Macro
'
    theSheet.Activate
    
    theSheet.Range("j2:j9").Select
    Range(Selection, Selection.Offset(0, facCount - 1)).Select
    Selection.FillRight
    
End Sub


' fill the formulas inside the table
Sub fillTableForumulas()
'
' fillTableForumulas Macro

    Dim oListObj As ListObject
    
    theSheet.Activate

    Set oListObj = theSheet.ListObjects(1)

    oListObj.ListColumns(10).DataBodyRange.Select
         
    Range(Selection, Selection.Offset(0, facCount - 1)).Select
    
    Selection.FillRight
End Sub



