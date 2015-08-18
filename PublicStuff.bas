Attribute VB_Name = "PublicStuff"
Public Rebuild As Boolean

Dim theSheet As Worksheet, oListObj As ListObject, Reason As Object, _
    resultRange As Range, columnCount As Variant


Public Sub NAConclusion()
    
    Application.ScreenUpdating = False
    
    ' make sure the worksheet activate event doesn't do it's whole thing
    Rebuild = True
    
    For Each wsSheet In ActiveWorkbook.Worksheets

    If Left(wsSheet.Name, 2) = "BP" Then
            
        Set oListObj = wsSheet.ListObjects(1)
        
        ' here is a maintenance item if columns are added to the table
        
        columnCount = oListObj.ListColumns.Count - 8
    
        For Each Reason In oListObj.ListColumns("Reason for Conclusion").DataBodyRange
            Debug.Print "Reason currently is " & Reason.Value
            
            ' don't overwrite if there the cell has content and _
                check if it says Not Applicable... and reset the value in that case
            
            If Reason.Value = "" Or _
               Reason.Value = "Not Applicable to all facilities in the property." Then
                
                Set resultRange = Range(Reason.Offset(0, 1), Reason.Offset(0, columnCount))
                ' resultRange.Activate
                Debug.Print WorksheetFunction.CountIf(resultRange, "N/A")
                
                If WorksheetFunction.CountIf(resultRange, "N/A") = columnCount Then
        
                    Reason.Value = "Not Applicable to all facilities in the property."
                
                Else
                    
                    Reason.Value = ""
                    
                End If
            
            End If
            
        Next Reason
    End If
    
    Next wsSheet

    
    Application.ScreenUpdating = True
    
    ' done, re-enable the worksheet activate event
    Rebuild = False


End Sub


Sub RebuildSummary()
    
    Dim sumSheet As Worksheet, sumList As ListObject, BPList As ListObject, _
        thesheets As Object, sumListRows As ListRows, sumListCols As ListColumns
        
     ' make sure the worksheet activate event doesn't do it's whole thing
    Rebuild = True
    
    Set sumSheet = ActiveWorkbook.Worksheets("Findings Summary")
         
    Set sumList = sumSheet.ListObjects("NCESummary")
    
    Set thesheets = ActiveWorkbook.Worksheets(Array("BP1 - Gas Exist Fac Des & Inst", _
        "BP2 - Gas New Fac Des & Inst", "BP3 - Gas Measurement", "BP4 - Gas Recording", _
        "BP5 - Gas Reporting", "BP6 - HC Liq Ex Fac Des & Inst", _
        "BP7 - HC Liq New Fac Des & Inst", "BP8 - HC Liquid Measurement", _
        "BP9 - HC Liquid Recording", "BP10 - HC Liquid Reporting", _
        "BP11 - Water Ex Fac Des & Inst", "BP12 - Water New Fac Des & Inst", _
        "BP13 - Water Measurement", "BP14 - Water Recording", "BP15 - Water Reporting"))

    Debug.Print sumList.ListColumns.Count
    
    Debug.Print thesheets.Item(1).ListObjects(1).ListColumns.Count
    
    Set sumListRows = sumList.ListRows
    Set sumListCols = sumList.ListColumns
    
    'delete the facility columns
    
    Debug.Print "Columns = "; sumListCols.Count
    
    Do While (sumListCols.Count > 11)
        
        sumListCols(12).Delete
    
    Loop
    
    
    'leave the first row to copy down formulas
          
    Debug.Print "Rows = " & sumListRows.Count
       
    
    Do While (sumListRows.Count > 1)
         
        sumListRows(2).Delete
    
    Loop
    
    ' add the facility columns based on the facility list
    
    ActiveWorkbook.Worksheets("Facility List").Range("FacIDs").Copy
    
    sumList.HeaderRowRange.End(xlToRight).Offset(0, 1).PasteSpecial (xlPasteValues)
    
    
    
    ' automate setting the facility count
    ' copy from the facility list sheet
    
    sumSheet.Range("FacIndex").Select
    sumSheet.Range(Range("FacIndex").Cells(2), Range("FacIndex").End(xlToRight)).Clear
    
    sumSheet.Range(Range("FacIndex").Cells(1), _
        sumList.HeaderRowRange.End(xlToRight).Offset(-1, 0)).DataSeries _
        Rowcol:=xlRows, Type:=xlLinear, Date:=xlDay, Step _
        :=1, Trend:=False
    

    
    ' sumList.ListColumns("Reporting Theme").DataBodyRange.Select
    
    For Each sht In thesheets
        
        Debug.Print "Copying data from " & sht.Name
        
        sht.ListObjects(1).DataBodyRange.Copy
        
        If Left(sht.Name, 4) = "BP1 " Then
            'first table pastes over remaining row to keep forumulas
            
            sumList.ListColumns("Reporting Theme").DataBodyRange.PasteSpecial _
                (xlPasteValues)
        
        Else
            'the rest paste on the next blank row at the bottom of the new table
            
            sumList.ListColumns("Reporting Theme").DataBodyRange.End(xlDown) _
             .Offset(1, 0).PasteSpecial (xlPasteValues)
        
        End If
        
    Next sht
    
    ' done, re-enable the worksheet activate event
    Rebuild = True

End Sub


