Attribute VB_Name = "AddDiscussionPoints"

Sub AddDiscPoints()
    Dim theComments As Range, comp As String, thePoint As String, searchRange As Range, foundIt As Range, _
    commentBox As ShapeRange
    
    
    Rebuild = True
    
    Application.ScreenUpdating = False
    
    ActiveWorkbook.Sheets("NCE Component").Activate
    
    Set thetable = ActiveWorkbook.Sheets("NCE Component").ListObjects("NCESub")
    'Set zoink = Range("NCESub[[#data],[Discussion Points]]")
    Set theComments = thetable.ListColumns("Discussion Points").DataBodyRange
        
    headerrow = thetable.HeaderRowRange.Row
    ' see if the NCE has a discussion point

    For Each aComment In theComments
        Debug.Print aComment; "  row "; aComment.Row
        If aComment.Value <> "" Then
        
            ' which NCE is it for
            ActiveWorkbook.Sheets("NCE Component").Activate

            aComment.Activate
            
            comp = thetable.ListColumns("NCE Component").DataBodyRange.Rows(aComment.Row - headerrow).Value
            
            
            thePoint = thetable.ListColumns("Discussion Points").DataBodyRange.Rows(aComment.Row - headerrow).Value
            
            Debug.Print comp; "  "; thePoint
            
            ' now add it to the BP table in the correct spot
            For Each wsSheet In ActiveWorkbook.Worksheets
    
                Debug.Print wsSheet.Name
        
                If Left(wsSheet.Name, 2) = "BP" Then
                    wsSheet.Activate
                    Set bptable = wsSheet.ListObjects(1)
                    firstrow = bptable.HeaderRowRange.Row
                    Set searchRange = bptable.ListColumns("NCE Component").DataBodyRange
                    Set foundIt = searchRange.Find(comp)
                    If Not foundIt Is Nothing Then
                        firstAddress = foundIt.Address
                        Do
                       
                            If Len(foundIt) = Len(comp) Then
                                'foundIt.Activate
                                bptable.ListColumns("NCE Component Description").DataBodyRange _
                                    .Rows(foundIt.Row - firstrow).Select
                                Debug.Print foundIt.Row
                                With Selection
                                    .ClearComments
                                    .addComment
                                    .Comment.Text Text:=thePoint
                                    .Comment.Visible = False
                                        
                                End With
                                
                            End If
                        Set foundIt = searchRange.FindNext(foundIt)
                        Loop While Not foundIt Is Nothing And _
                            foundIt.Address <> firstAddress
                    End If
                
                End If
                

            Next wsSheet
        
        End If
    Next aComment
    
    Application.ScreenUpdating = True
        
    Rebuild = False
    
End Sub

Sub scalecomments()
' scale all the comment boxes
    For Each wsSheet In ActiveWorkbook.Worksheets

        Debug.Print wsSheet.Name

    
        If Left(wsSheet.Name, 4) = "BP1 " Then
            wsSheet.Activate
    
            For Each cbox In wsSheet.Shapes
                If Left(cbox.Name, 4) = "Comm" Then
                    ' Debug.Print "name "; CBox.Name
                    ' Debug.Print "h "; CBox.Height
                    ' Debug.Print " width "; CBox.Width
                    ' Debug.Print "title "; " txt "; CBox.AlternativeText
                    
                    ' reset height to default
                    cbox.Height = 59.25
                    
                    cLength = Len(cbox.AlternativeText)
                    
                    ' Debug.Print "length "; cLength
                    
                    ' comment boxes are 5 lines of ~22 = 110
                    If cLength > 110 Then
                        cScale = WorksheetFunction.RoundUp((cLength / 20) / 5, 1)
                        With cbox
                            .ScaleHeight cScale, msoFalse
                            .Top = 50
                            .Left = 50
                            .Width = 500
                            .Height = 100
                        End With
                        With cbox.TextFrame.Characters.Font
                            .Name = "Verdana"
                            .Size = 12
                        End With

                    End If
                End If
            Next cbox
        End If
                
    Next wsSheet

End Sub

Sub FormatAllComments()
'www.contextures.com/xlcomments03.html
  Dim ws As Worksheet
  Dim cmt As Comment
  For Each wsSheet In ActiveWorkbook.Worksheets
    If Left(wsSheet.Name, 4) = "BP1 " Then
        wsSheet.Activate
    
        For Each cmt In wsSheet.Comments
          With cmt.Shape.TextFrame.Characters.Font
            .Name = "Verdana"
            .Size = 12
          End With
          With cmt.Shape
            .Top = 50
            .Left = 50
            .Width = 500
            .Height = 100
          End With
        Next cmt
    End If
  Next wsSheet
End Sub

