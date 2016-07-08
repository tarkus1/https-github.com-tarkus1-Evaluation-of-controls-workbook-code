VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet22"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim tbl As ListObject, ncetbl As ListObject, firstrow As Integer, looktxt As String, nceRng As Range, _
    cmText As String, cmLength As Variant, cmHeight As Variant
    
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
                Set ncetbl = Worksheets("NCE Component").ListObjects(1)
                Set nceRng = Worksheets("NCE Component").Range("B2", "k236")
                
                ' Set nceRng1 = Range("NCESub")
                ' Debug.Print ActiveCell.Address
                ' Debug.Print firstrow
                ' Debug.Print Selection.Row
                ' Debug.Print nceRng.Address
                
                ' Debug.Print Range(ncetbl.ListColumns("NCE").DataBodyRange, ncetbl.ListColumns("Discussion Points").DataBodyRange).Address
                
        
                
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
                    .Top = 50
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
