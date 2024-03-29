Attribute VB_Name = "WorkBookUtils"
Sub CheckContent()
    Dim objlist As Object, found As Object, wsSheet As Worksheet
    
    Rebuild = True

    For Each wsSheet In ActiveWorkbook.Worksheets

        If Left(wsSheet.Name, 2) = "BP" Then
        
            wsSheet.Activate
            
            Set objlist = wsSheet.ListObjects(1)
            
            With objlist.ListColumns("Conclusion").DataBodyRange
                Set found = .Find("N/A")
                If Not found Is Nothing Then
                    firstAddress = found.Address
                    Do
                        found.Select
                        If found.Offset(0, 1) = "" _
                           Or found.Offset(0, 2) = "" _
                           Or found.Offset(0, 2) = "" _
                           Or found.Offset(0, 3) = "" Then
                            MsgBox "Missing content"
                            Stop
                        End If
                           
                        If found.Offset(0, 2) <> "N/A" Then MsgBox "found incorrect performer"
                        Set found = .FindNext(found)
                    Loop While Not found Is Nothing And found.Address <> firstAddress
                End If
            End With
            Debug.Print "complete "; ActiveSheet.Name
            
        End If
    Next wsSheet
    
    Rebuild = False

End Sub

Sub CreatePDF()
'
' CreatePDF Macro
'
    Dim wBook As Workbook, theName As Variant, thePath As Variant

'
    Rebuild = True
    
    Set Workbook = ActiveWorkbook
    
    thePath = Workbook.Path
    Debug.Print thePath
    
    theName = Left(Workbook.Name, Len(Workbook.Name) - 5)
    Debug.Print thePath & "/" & theName
    
    
        Sheets(Array("BP1 - Gas Exist Fac Des & Inst", "BP2 - Gas New Fac Des & Inst", _
        "BP3 - Gas Measurement", "BP4 - Gas Recording", "BP5 - Gas Reporting", _
        "BP6 - HC Liq Ex Fac Des & Inst", "BP7 - HC Liq New Fac Des & Inst", _
        "BP8 - HC Liquid Measurement", "BP9 - HC Liquid Recording", _
        "BP10 - HC Liquid Reporting", "BP11 - Water Ex Fac Des & Inst", _
        "BP12 - Water New Fac Des & Inst", "BP13 - Water Measurement", _
        "BP14 - Water Recording", "BP15 - Water Reporting")).Select
    Sheets("BP1 - Gas Exist Fac Des & Inst").Activate
    Range("I2:I9").Select
    With Selection.Font
        .Name = "Arial"
        .FontStyle = "Regular"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    
    ' Old style facility list
   ' Workbook.Sheets(Array("Handout", "Facility List", "BP1 - Gas Exist Fac Des & Inst", _
        "BP2 - Gas New Fac Des & Inst", "BP3 - Gas Measurement", "BP4 - Gas Recording", _
        "BP5 - Gas Reporting", "BP6 - HC Liq Ex Fac Des & Inst", _
        "BP7 - HC Liq New Fac Des & Inst", "BP8 - HC Liquid Measurement", _
        "BP9 - HC Liquid Recording", "BP10 - HC Liquid Reporting", _
        "BP11 - Water Ex Fac Des & Inst", "BP12 - Water New Fac Des & Inst", _
        "BP13 - Water Measurement", "BP14 - Water Recording", "BP15 - Water Reporting")). _
        Select
        
    ' new tabular facility list
    
    Workbook.Sheets(Array("Handout", "Tabular Facility List", "BP1 - Gas Exist Fac Des & Inst", _
        "BP2 - Gas New Fac Des & Inst", "BP3 - Gas Measurement", "BP4 - Gas Recording", _
        "BP5 - Gas Reporting", "BP6 - HC Liq Ex Fac Des & Inst", _
        "BP7 - HC Liq New Fac Des & Inst", "BP8 - HC Liquid Measurement", _
        "BP9 - HC Liquid Recording", "BP10 - HC Liquid Reporting", _
        "BP11 - Water Ex Fac Des & Inst", "BP12 - Water New Fac Des & Inst", _
        "BP13 - Water Measurement", "BP14 - Water Recording", "BP15 - Water Reporting")). _
        Select
        
    
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
    
   ' ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, filename:=thePath & "/" & theName _
        , Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
        :=False, OpenAfterPublish:=True
        
        Sheets(Array("BP1 - Gas Exist Fac Des & Inst", "BP2 - Gas New Fac Des & Inst", _
        "BP3 - Gas Measurement", "BP4 - Gas Recording", "BP5 - Gas Reporting", _
        "BP6 - HC Liq Ex Fac Des & Inst", "BP7 - HC Liq New Fac Des & Inst", _
        "BP8 - HC Liquid Measurement", "BP9 - HC Liquid Recording", _
        "BP10 - HC Liquid Reporting", "BP11 - Water Ex Fac Des & Inst", _
        "BP12 - Water New Fac Des & Inst", "BP13 - Water Measurement", _
        "BP14 - Water Recording", "BP15 - Water Reporting")).Select
    Sheets("BP15 - Water Reporting").Activate
    ActiveWindow.SmallScroll Down:=-24
    Range("I2:I9").Select
    With Selection.Font
        .Name = "Arial"
        .FontStyle = "Regular"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    Sheets("BP1 - Gas Exist Fac Des & Inst").Select
    
    Rebuild = False
    
End Sub

Sub shiftRight()
'
' shiftRight Macro
'

'
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
End Sub


Sub concFormat()
'
' concFormat Macro
'

'
    'ActiveWorkbook.ActiveSheet.ListObjects(1).ListColumns("Reason for Conclusion").DataBodyRange.Select
    
    Dim objlist As Object, found As Object, wsSheet As Worksheet
    
    Rebuild = True

    For Each wsSheet In ActiveWorkbook.Worksheets

        If Left(wsSheet.Name, 2) = "BP" Then
        
            wsSheet.Activate
            
            Set objlist = wsSheet.ListObjects(1)
            
            objlist.ListColumns("Reason for Conclusion").DataBodyRange.Select
            
            With Selection
                .HorizontalAlignment = xlLeft
                .VerticalAlignment = xlTop
                .WrapText = True
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
            End With
            With Selection.Font
                .Name = "Arial"
                .Size = 12
                .Strikethrough = False
                .Superscript = False
                .Subscript = False
                .OutlineFont = False
                .Shadow = False
                .Underline = xlUnderlineStyleNone
                .ColorIndex = xlAutomatic
                .TintAndShade = 0
                .ThemeFont = xlThemeFontNone
            End With
            With Selection.Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            
        End If
    Next wsSheet
    
    Rebuild = False
            
End Sub

Sub rowHeights()
    
    ' autofit all rows in the BP sheets and make sure they are _
        at least 30 high
        
    Dim objlist As Object, found As Object, wsSheet As Worksheet
    
    Rebuild = True

    For Each wsSheet In ActiveWorkbook.Worksheets

        If Left(wsSheet.Name, 2) = "BP" Then
        
            wsSheet.Activate
            
            Set objlist = wsSheet.ListObjects(1)
            
            objlist.ListColumns("Conclusion").DataBodyRange.Rows.EntireRow.AutoFit
            
            For Each Rw In objlist.ListColumns("Conclusion").DataBodyRange.Rows
            
                If Rw.rowHeight < 30 Then Rw.rowHeight = 30
                
            
            Next Rw
            
        End If
    Next wsSheet
    
    Rebuild = False

End Sub

Sub addObs()
    ' add observation label below last row
        
    Dim objlist As Object, found As Object, wsSheet As Worksheet, tname As Range, tt As String
    
    
    Rebuild = True

    For Each wsSheet In ActiveWorkbook.Worksheets

        If Left(wsSheet.Name, 2) = "BP" Then
        
            wsSheet.Activate
            
            Set objlist = wsSheet.ListObjects(1)
            
            tt = objlist.Name & "[[#Headers],[NCE Component]]"
            
            
            Set tname = Range(tt)
            
            Debug.Print tname
            
            
            tname.End(xlDown).Offset(2, 0).Select
            
            Selection.Activate
            
            
            Selection.Value = "Observations:"
            
                                  
        End If
   Next wsSheet
    
    Rebuild = False



End Sub
