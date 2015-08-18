Attribute VB_Name = "findingsSummary"
Dim facSumSh As Worksheet, findSumSh As Worksheet, findObj As ListObject, _
    facObj As ListObject, findSumCols As ListColumns, NCERange As Range, _
    RiskRange As Range, facID As Variant

Sub RefreshFindingsSummary()
    Set facSumSh = ActiveWorkbook.Worksheets("TestSummary")
    Set findSumSh = ActiveWorkbook.Worksheets("Findings Summary")

    Set findObj = findSumSh.ListObjects("NCESummary")
    
    Set facObj = facSumSh.ListObjects(1)
    
    Set findSumCols = findObj.ListColumns

    Set NCERange = Range(findSumCols(3).DataBodyRange, _
        findSumCols(5).DataBodyRange)
        
    Set RiskRange = findObj.ListColumns("NCE Risk").DataBodyRange
    
    'NCERange.Select
    'RiskRange.Select
    
    facObj.DataBodyRange.Delete
    
    'play with this!!!!!
    Debug.Print facObj.ListRows.Row.Index
    
    
    For Each fcol In findSumCols
    
        Debug.Print fcol.Name
        
        If UCase(Left(fcol.Name, 2)) = "AB" Then
            
            facID = fcol.Name
            
            Debug.Print "Copy data for "; fcol.Name
            
            fcol.DataBodyRange.Copy
            
            facSumSh.Range("TestSummary[Conclusion]").ListRows.Item(0).PasteSpecial (xlPasteValues)
            
            facSumSh.Range("TestSummary[Facility]") = facID
            
            NCERange.Copy
            
            facSumSh.Range("TestSummary[Reporting Theme]").PasteSpecial (xlPasteValues)
            
        
        End If
    
    Next fcol

End Sub
