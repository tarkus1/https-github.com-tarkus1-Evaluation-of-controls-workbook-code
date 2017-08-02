Attribute VB_Name = "findingsSummary"
Dim facSumSh As Worksheet, findSumSh As Worksheet, findObj As ListObject, _
    facObj As ListObject, findSumCols As ListColumns, NCERange As Range, _
    RiskRange As Range, facID As Variant, facIndex As Variant, _
    facNum As Variant, curRows As Variant, newRows As Variant

Sub RefreshFindingsSummary()

    ' make sure the worksheet activate event doesn't do it's whole thing
    
    Rebuild = True
    
    
    Set facSumSh = ActiveWorkbook.Worksheets("Findings Summary by Facility")
    Set findSumSh = ActiveWorkbook.Worksheets("Findings Summary")

    Set findObj = findSumSh.ListObjects("NCESummary")
    
    Set facObj = facSumSh.ListObjects(1)
    
    Set findSumCols = findObj.ListColumns

    Set NCERange = Range(findSumCols(3).DataBodyRange, _
        findSumCols(5).DataBodyRange)
        
    Set RiskRange = findObj.ListColumns("NCE Risk").DataBodyRange
    
   
    'facObj.DataBodyRange.Delete
    
    facNum = 1
    
    
    For Each fcol In findSumCols
    
        Debug.Print fcol.Name
        
        
        If UCase(Left(fcol.Name, 2)) = "AB" Or UCase(Left(fcol.Name, 2)) = "SK" Then
            
            facID = fcol.Name
            
            facIndex = fcol.Index
            
            Debug.Print "Copy data for "; fcol.Name; fcol.Index
            
            If facNum = 1 Then
            
                  fcol.DataBodyRange.Copy
                  
                  facSumSh.Range("FacSumm[Conclusion]").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                      :=False, Transpose:=False
                 
                  facSumSh.Range("FacSumm[Facility]") = facID
                  
                  facSumSh.Range("FacSumm[Facility Number]") = facNum
                  
                  NCERange.Copy
                  
                             
                  facSumSh.Range("FacSumm[Reporting Theme]").PasteSpecial (xlPasteValues)
                  
                  RiskRange.Copy
                  
                  facSumSh.Range("FacSumm[NCE Risk]").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                      :=False, Transpose:=False
                    
                  facNum = facNum + 1
                    
            Else
                Debug.Print "facility number "; facNum
                
                curRows = facObj.ListRows.Count
                
                fcol.DataBodyRange.Copy
                
                facObj.ListColumns("Conclusion").DataBodyRange.End(xlDown) _
                    .Offset(1, 0).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, _
                    SkipBlanks:=False, Transpose:=False
                    
                newRows = facObj.ListRows.Count
                    
                Range(facObj.ListColumns("Facility").DataBodyRange.Rows(curRows).Offset(1, 0), _
                    facObj.ListColumns("Facility").DataBodyRange.Rows(newRows)) = facID
                    
                Range(facObj.ListColumns("Facility Number").DataBodyRange.Rows(curRows).Offset(1, 0), _
                    facObj.ListColumns("Facility Number").DataBodyRange.Rows(newRows)) = facNum
                    
                NCERange.Copy
          
                facObj.ListColumns("Reporting Theme").DataBodyRange.Rows(curRows).Offset(1, 0).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                      :=False, Transpose:=False
                
                RiskRange.Copy
                
                facObj.ListColumns("NCE Risk").DataBodyRange.Rows(curRows).Offset(1, 0).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                      :=False, Transpose:=False
                
                facNum = facNum + 1
            
            End If
        End If
    
    Next fcol
    
    facSumSh.PivotTables("PivotTable1").RefreshTable
    
    facSumSh.Range("a1").Select

    
    ' done, re-enable the worksheet activate event

    Rebuild = False
    
End Sub
