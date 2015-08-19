Attribute VB_Name = "Module5"
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    ActiveWorkbook.Worksheets("TestSummary").Range("testrange").Copy
    
    x = ActiveWorkbook.Worksheets("TestSummary").ListObjects("TestSummary").ListRows.Count
    
    ActiveWorkbook.Worksheets("TestSummary").ListObjects("TestSummary").ListColumns("Facility").DataBodyRange.Rows(x).Offset(1, 0).Select
        
        
    For Each tcol In ActiveWorkbook.Worksheets("TestSummary").ListObjects("TestSummary").ListColumns
    
        Debug.Print tcol.Index; "  "; tcol.Name
    
    Next tcol
    
    ActiveWorkbook.Worksheets("TestSummary").ListObjects("TestSummary").ListColumns("Facility") _
        .DataBodyRange.End(xlDown).Offset(1, 0).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    ActiveWorkbook.Worksheets("TestSummary").ListObjects("TestSummary").ListColumns(2) _
        .DataBodyRange.End(xlDown).Offset(1, 0).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
        
    
    End Sub
Sub Macro5()
Attribute Macro5.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro5 Macro
'

'
    Range("TestSummary[[#Headers],[Facility]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("C6").Select
    Selection.End(xlDown).Select
    Range("C15").Select
End Sub
