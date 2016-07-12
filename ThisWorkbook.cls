VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Dim TargetSheet As Worksheet, SourceSheet As Worksheet, _
    TargetName As String, SourceName As String, tNCE As String, sConc As String, _
    TargetNum As Variant, SourceNum As Variant, _
    SoListObj As ListObject, TgListObj As ListObject, _
    Srow As Variant, TRow As Variant, firstrow As Variant, temp As Variant
     
    
Public Sub Workbook_SheetActivate(ByVal Sh As Object)

    
    
    ' MsgBox Sh.Name
    
'   watch for recursion when the worksheet is activated within this routine _
    and when activated during rebuilding
    
    If SourceSheet Is Nothing And Left(Sh.Name, 2) = "BP" And Not Rebuild _
        And Not Left(Sh.Name, 3) = "BPT" Then
        
        Set TargetSheet = Sh
        
        TargetName = Left(TargetSheet.Name, 4)
        TargetNum = Mid(TargetSheet.Name, 3, 2)
        ' MsgBox TargetNum
        
        If TargetNum > 5 Then
            SourceNum = TargetNum - 5
            ' MsgBox SourceNum
            
            SourceName = "BP" & SourceNum
            ' MsgBox Len(SourceName)
            
            Set TgListObj = TargetSheet.ListObjects(1)
            
            
            For Each wsSheet In ActiveWorkbook.Worksheets
                Debug.Print Len(WorksheetFunction.Trim(Left(wsSheet.Name, 4)))
                If WorksheetFunction.Trim(Left(wsSheet.Name, 4)) = SourceName Then
                    Set SourceSheet = wsSheet
                    Set SoListObj = SourceSheet.ListObjects(1)
                                       
                    ' do the grab of stuff here
                    firstrow = TgListObj.ListColumns("Conclusion").DataBodyRange.Row - 1
                    
                    For Each tconc In TgListObj.ListColumns("Conclusion").DataBodyRange.Cells
                        
                        TRow = tconc.Row - firstrow

                        If tconc.Value = Empty And TgListObj.ListColumns("Control Performer Role"). _
                            DataBodyRange.Item(TRow).Value = Empty Then
                            
                            tNCE = TgListObj.ListColumns("NCE").DataBodyRange.Item(TRow).Value
                            
                            For Each sNCE In SoListObj.ListColumns("NCE").DataBodyRange.Cells
                                If sNCE.Value = tNCE Then
                                
                                    Srow = sNCE.Row - firstrow
                                    
                                    'Conclusion
                                    TgListObj.ListColumns("Conclusion").DataBodyRange.Cells(TRow).Value = _
                                    SoListObj.ListColumns("Conclusion").DataBodyRange.Cells(Srow).Value
                                   
                                    'Evidence
                                    TgListObj.ListColumns("Evidence").DataBodyRange.Cells(TRow).Value = _
                                    SoListObj.ListColumns("Evidence").DataBodyRange.Cells(Srow).Value
                                    
                                    'Control Performer
                                    TgListObj.ListColumns("Control Performer Role").DataBodyRange.Cells(TRow).Value = _
                                    SoListObj.ListColumns("Control Performer Role").DataBodyRange.Cells(Srow).Value
                                   
                                    'Reason -- don't copy formulas
                                    If Left(SoListObj.ListColumns("Reason for Conclusion").DataBodyRange.Cells(Srow).Formula, 3) <> "=IF" Then
                                        TgListObj.ListColumns("Reason for Conclusion").DataBodyRange.Cells(TRow).Value = _
                                        SoListObj.ListColumns("Reason for Conclusion").DataBodyRange.Cells(Srow).Value
                                    End If
                                    
                                End If
                                
                              '  MsgBox snce.Row
                            Next sNCE
                        
                        End If
                    Next tconc
                    
                End If
            Next wsSheet
            
            
        End If
    End If
    Set SourceSheet = Nothing
End Sub






