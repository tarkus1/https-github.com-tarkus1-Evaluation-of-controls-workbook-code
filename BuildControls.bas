Attribute VB_Name = "BuildControls"

Dim wsSheets As Worksheets, wsSource As Worksheet, wsSheet As Worksheet, _
    wsNCE As Worksheet, wsClient As Worksheet, oListObj As ListObject, _
    headerRange As Range, BPNum As Integer

Sub BuildControls()

    ' make sure the worksheet activate event doesn't do it's whole thing
    Rebuild = True
    
    Application.ScreenUpdating = False
    
    For Each wsSheet In ActiveWorkbook.Worksheets
    
        Debug.Print wsSheet.Name
        
        If Left(wsSheet.Name, 2) = "BP" Then
            
                Set theSheet = wsSheet
                BPNum = WorksheetFunction.Trim(Mid(wsSheet.Name, 3, 2))
                Debug.Print "working on " & theSheet.Name
                
                DeleteTableRows
                FilterBPNCEs
                'FilterClientControls
                CopyToTable
                CorrectFormatting
                ResetPrintArea
                
                ' RebuildSummary
    
            
        End If
    Next wsSheet
    
    ' done, re-enable the worksheet activate event
    Rebuild = False
    Application.ScreenUpdating = True

End Sub

Sub DeleteTableRows()
    Dim TRow As Object, firstrow As Variant, oListObj As Object, _
        oListRows As ListRows, RowIndex As Variant
    
    'Set wsSheet = ActiveWorkbook.Worksheets("BPT3 - Gas MeasurementTest")
    

    Debug.Print "in DeleteTableRows for the " & wsSheet.Name
    
    Set oListObj = wsSheet.ListObjects(1)
    Set oListRows = oListObj.ListRows
    
    'leave the first row to copy down formulas
          
    Debug.Print oListRows.Count
    
    Do While (oListRows.Count > 1)
         
        oListRows(2).Delete
    
    Loop

End Sub

Sub FilterBPNCEs()

    ' Set wsSheet = ActiveWorkbook.Worksheets("BP1 - Gas Exist Fac Des & Inst")

    Set wsNCE = ActiveWorkbook.Worksheets("NCE Component")
    Debug.Print BPNum
    
    Debug.Print "in FilterBPNCEs for the " & wsSheet.Name
        
    Range("NCE_BP") = BPNum

    wsNCE.Range("NCESub[#All]").AdvancedFilter Action:=xlFilterCopy, CriteriaRange:= _
        wsNCE.Range("q1:q2"), CopyToRange:=wsNCE.Range("s1:y1"), Unique:=True

End Sub
            
Sub FilterClientControls()
    Dim newCriteria As Range
    
    ' Set wsSheet = ActiveWorkbook.Worksheets("BP1 - Gas Exist Fac Des & Inst")

    Set wsClient = ActiveWorkbook.Worksheets("Client Controls")
    Set wsNCE = ActiveWorkbook.Worksheets("NCE Component")
    Debug.Print "in FilterClientControls for the " & wsSheet.Name
    
    ' clear the filter criteria
    
    wsClient.Range(Range("CC_Criteria").Offset(1, 0), Range("CC_Criteria").End(xlDown)).Clear

    
    wsNCE.ListObjects.Add(xlSrcRange, wsNCE.Range("Extract").CurrentRegion, , xlYes).Name = _
        "BP_NCEs"
        
    wsNCE.ListObjects("BP_NCEs").ListColumns("NCEProd").DataBodyRange.Copy
    
    wsClient.Range("CC_Criteria").Offset(1, 0).PasteSpecial (xlPasteValues)
    
    wsClient.Range("CC_Criteria")(1, 1).Offset(1, 1) = BPNum
    
    
    wsNCE.ListObjects("BP_NCEs").Unlist
    
    ' reset the criteria range
    
    Set newCriteria = wsClient.Range(Range("CC_Criteria"), Range("CC_Criteria").End(xlDown))
    
    ' need to clear the data from the table?
    ' wsClient.Range(Range("Extract").Offset(1, 0), Range("Extract").End(xlDown)).Select

    newCriteria(1, 1).Offset(0, 1).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(0, 1).Select
    Range(Selection, Selection.End(xlUp)).Offset(1, 0).Select
    Selection.FillDown
    
    wsClient.Range("ClientControls[#All]").AdvancedFilter Action:=xlFilterCopy, _
        CriteriaRange:=newCriteria, CopyToRange:=wsClient.Range( _
        "Extract"), Unique:=True
    

End Sub
            
Sub CopyToTable()
    
    Set wsClient = ActiveWorkbook.Worksheets("Client Controls")
    Set wsNCE = ActiveWorkbook.Worksheets("NCE Component")

    Debug.Print "in CopyToTable for the " & wsSheet.Name
    'copy NCE proxy controls
    
    wsNCE.ListObjects.Add(xlSrcRange, wsNCE.Range("Extract").CurrentRegion, , xlYes).Name _
        = "BP_NCEs"
    
    wsNCE.Range("BP_NCEs[[Theme]:[NCE Component Description1]]").Copy
    
    wsSheet.ListObjects(1).ListColumns("Theme").DataBodyRange.PasteSpecial _
        Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    ' now grab any comments before the table gets toasted
    
    ' see module1 for work in progress
        
    wsNCE.ListObjects("BP_NCEs").Unlist
    
    'copy client controls at end of table
    
    wsClient.ListObjects.Add(xlSrcRange, wsClient.Range("Extract").CurrentRegion, , xlYes).Name _
        = "BP_ClientControls"
    
    wsClient.Range("BP_ClientControls[[Theme]:[Client Control Description]]").Copy
    
    wsClient.ListObjects("BP_ClientControls").Unlist
    
    wsSheet.ListObjects(1).ListColumns("Theme").DataBodyRange.End(xlDown).Offset(1, 0).PasteSpecial _
        Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    

End Sub
            
Sub CorrectFormatting()

    ' Set wsSheet = ActiveWorkbook.ActiveSheet
    Set oListObj = wsSheet.ListObjects(1)
    Debug.Print oListObj
    
    oListObj.Sort.SortFields.Clear
    oListObj.Sort.SortFields.Add Key:=oListObj.ListColumns("Theme").DataBodyRange, SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    oListObj.Sort.SortFields.Add Key:=oListObj.ListColumns("NCE").DataBodyRange, SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With oListObj.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    'set the row height
    
    oListObj.ListColumns("NCE Component Description").DataBodyRange.Rows.EntireRow.AutoFit
                        
    For Each Rw In oListObj.ListColumns("NCE Component Description").DataBodyRange.Rows
    
        If Rw.rowHeight < 30 Then Rw.rowHeight = 30
            
    Next Rw

    
    'correct text alignment for NCE Component
    
    With oListObj.ListColumns("NCE Component").Range
        .ClearFormats
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = True

    End With
    

End Sub

            
Sub ResetPrintArea()
    Dim BottomRight As Object, PrArea As Range
    ' Set wsSheet = ActiveWorkbook.ActiveSheet

    Debug.Print "in ResetPrintArea for the " & wsSheet.Name
    
    ' make sure the Reason column is blank
    wsSheet.ListObjects(1).ListColumns("Reason for Conclusion").DataBodyRange.Clear
    
    ' for the lower right, select 3 rows past the bottom of Reason for conclusion
    Set BottomRight = wsSheet.ListObjects(1).ListColumns("Reason for Conclusion").DataBodyRange.End(xlDown).Offset(3, 0)
    
    Set PrArea = wsSheet.Range(("C1"), BottomRight)
    
    wsSheet.PageSetup.PrintArea = PrArea.Address
    
End Sub


Sub ColumnWidths()
'
' ColumnWidths Macro
'

'
    Columns("D:D").Select
    Selection.ColumnWidth = 75
    Columns("e:e").Select
    Selection.ColumnWidth = 6
    Columns("g:h").Select
    Selection.ColumnWidth = 12
    Columns("f:f").Select
    Selection.ColumnWidth = 16
    Columns("i:i").Select
    Selection.ColumnWidth = 26
    Columns("A:C").Select
    Selection.ColumnWidth = 8.43
    
End Sub
