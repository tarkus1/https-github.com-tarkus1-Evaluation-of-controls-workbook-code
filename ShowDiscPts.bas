Attribute VB_Name = "ShowDiscPts"
    Public tbl As ListObject, ncetbl As ListObject, firstrow As Integer, looktxt As String, nceRng As Range, _
        cmText As String, cmLength As Variant, cmHeight As Variant, visrng As Range


Public Sub showDiscussionPoint()
        
        If Not Rebuild Then
            
            For Each cmt In ActiveSheet.Comments
              cmt.Delete
            Next cmt
            
            Set tbl = Selection.ListObject
            If Not tbl Is Nothing Then
                
                firstrow = tbl.DataBodyRange.Row
                
                If tbl.HeaderRowRange.Cells(1, Selection.Column) _
                    = "NCE Component Description" And Selection.Row >= firstrow Then
                        
                    Debug.Print tbl.HeaderRowRange.Cells(1, Selection.Column)
                    Set visrng = ActiveWindow.VisibleRange
                    Set ncetbl = Worksheets("NCE Component").ListObjects(1)
                    Set nceRng = Worksheets("NCE Component").Range("B2", "k236")
                    
                    
                    looktxt = tbl.ListColumns("NCE").DataBodyRange.Rows(Selection.Row - firstrow + 1).Value
                    
                    Debug.Print looktxt
                    
                    cmText = Application.WorksheetFunction.VLookup(looktxt, nceRng, 10, False)
                    
                    cmLength = Len(cmText)
                    Debug.Print cmLength
                    cmHeight = Application.WorksheetFunction.RoundUp(cmLength / 65, 0) * 17
                    Debug.Print cmHeight
                    
                    Set curcmt = ActiveCell.addComment
                    
                    curcmt.Text Text:=cmText
                    
                    With curcmt.Shape
                        .Top = visrng.Top + 50
                        .Left = 50
                        .Width = 480
                        .Height = cmHeight
                        
                    End With
                    
                    With curcmt.Shape.TextFrame.Characters.Font
                        .Name = "Verdana"
                        .Size = 12
                        
                    End With
                    
                    curcmt.Visible = True
                End If
            End If
        End If
    
End Sub
