Attribute VB_Name = "ApplicVsEval"

Sub ApplicableVsEvaluated()

'
'   look at the client's portfolio of facilities to determine if any are applicable to each theme _
    and then set whether the theme can be "not applicable" or "not evaluated" _
    reference: https://aer.ca/documents/projects/epap/EPAP_AvailableOptionsforDeclarationThemeConclusion.pdf
'


Dim themeObj As Object, wsSheet As Worksheet, firstrow As Variant, curCol As Variant, _
    subTypeList() As Variant, arrIndx As Integer, arrDim As Integer, foundOne As Boolean, _
    thmNum As Integer

Set wsSheet = ActiveWorkbook.Worksheets("Theme Applicability")
Set themeObj = wsSheet.ListObjects("ThemeApplic")

firstrow = themeObj.HeaderRowRange.Row

Debug.Print "First row "; firstrow

For Each thm In themeObj.ListRows

    If Val(thm.Range.Item(1)) >= 1 Then

        thmNum = thm.Range.Item(1).Value
                            
        Debug.Print thmNum
        
        arrIndx = 1

        arrDim = Application.WorksheetFunction.CountIf(thm.Range, "Y")
        
        ReDim subTypeList(1 To arrDim)
        
        For Each thmtype In thm.Range
            
            
            
            curCol = thmtype.Column - 1
            
            Debug.Print thmtype.Value; curCol + 1
            
            If UCase(thmtype.Value) = "Y" Then _
            
                Debug.Print themeObj.HeaderRowRange.Item(curCol + 1).Value     '("SubType").DataBodyRange.Item(curRow).Value

                ' work here on indx and increment
                
                subTypeList(arrIndx) = themeObj.HeaderRowRange.Item(curCol + 1).Value
                
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
                Debug.Print "Theme "; thm.Range.Item(1).Value; " found subtype "; stype
            End If
        
        Next stype
        
        If foundOne = True Then
            ' client has an applicable facility
            
            Range("ThemeApplic[Evaluated]").Item(thmNum).Value = "Not Evaluated"
        
        Else
            ' client has no applicable facilities
            
            Range("ThemeApplic[Evaluated]").Item(thmNum).Value = "Not Applicable"
            
        End If

    
    End If

Next thm

End Sub






