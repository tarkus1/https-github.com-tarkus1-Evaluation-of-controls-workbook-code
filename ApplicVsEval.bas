Attribute VB_Name = "ApplicVsEval"
Sub ApplicEval()

Dim themeObj As Object, wsSheet As Worksheet, firstrow As Variant, curRow As Variant, _
    subTypeList() As Variant, arrIndx As Integer, arrDim As Integer, foundOne As Boolean

Set wsSheet = ActiveWorkbook.Worksheets("Theme Applicability")
Set themeObj = wsSheet.ListObjects("ThemeSubTypes")

firstrow = themeObj.HeaderRowRange.Row

Debug.Print "First row "; firstrow

For Each thm In themeObj.ListColumns

    If Val(thm.Name) >= 1 Then
        Debug.Print thm.Name
                            
        arrIndx = 1

        arrDim = Application.WorksheetFunction.CountIf(thm.DataBodyRange, "Y")
        
        ReDim subTypeList(1 To arrDim)
        
        For Each thmtype In thm.DataBodyRange
            
            
            
            curRow = thmtype.Row - firstrow
            
            Debug.Print thmtype.Value; curRow
            
            If UCase(thmtype.Value) = "Y" Then _
            
                Debug.Print themeObj.ListColumns("SubType").DataBodyRange.Item(curRow).Value

                ' work here on indx and increment
                
                subTypeList(arrIndx) = themeObj.ListColumns("SubType").DataBodyRange.Item(curRow).Value
                
                Debug.Print "array "; arrIndx; subTypeList(arrIndx)
                
                arrIndx = arrIndx + 1
                
            End If
            
            'Debug.Print " subtype "; themeObj.ListColumns("SubType").ListRows.Index(1).Name
            
            
        Next thmtype
    
        foundOne = False

        For Each stype In subTypeList
        
            Debug.Print "array result "; stype
            
            If Not Range("petrinex[Facility Sub-Type]").Find(stype) Is Nothing Then
                foundOne = True
                Debug.Print "Theme "; thm.Name; " found subtype "; stype
            End If
        
        Next stype
        
        If foundOne = True Then
            thm.Range.Rows.Item(1).Offset(-1, 0).Value = "Not Applicable"
        
        Else
            thm.Range.Rows.Item(1).Offset(-1, 0).Value = "Not Evaluated"
        End If

    
    End If

Next thm

End Sub
