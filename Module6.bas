Attribute VB_Name = "Module6"
Sub nceRng()
Attribute nceRng.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ncerng Macro
'

'
    ActiveWindow.SmallScroll ToRight:=-2
    Range("NCESub[[#Headers],[NCE]]").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
End Sub
Sub Macro5()
Attribute Macro5.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro5 Macro
'

'
    Debug.Print Worksheets("NCE Component").ListObjects("NCESub").DataBodyRange.Row
    
    'Worksheets("NCE Component").ListObjects("NCESub").DataBodyRange.Select
    Debug.Print Range("NCESub[[#data],[NCE]:[Discussion Points]]").Rows(2).Address
    
    
End Sub
Sub Macro6()
Attribute Macro6.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro6 Macro
'

'
    Debug.Print Range("NCESub[[#data],[NCE Component]:[NCE Risk]]").Address
    ActiveWorkbook.Save
    Sheets("BP2 - Gas New Fac Des & Inst").Select
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    Sheets("BP1 - Gas Exist Fac Des & Inst").Select
    Range("D16").Select
    ActiveWorkbook.Save
    Range("D15").Select
    Application.Goto Reference:="NCESub"
    ActiveWindow.SmallScroll Down:=-15
    Range("B2").Select
    Sheets("BP1 - Gas Exist Fac Des & Inst").Select
    Range("D13").Select
    Sheets("NCE Component").Select
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 8
    Range("K10").Select
    Selection.End(xlDown).Select
    Sheets("BP1 - Gas Exist Fac Des & Inst").Select
    Range("D11").Select
    ActiveWorkbook.Save
    ActiveWorkbook.Save
    Range("D16").Select
    ActiveWindow.ActivateNext
    Sheets("BP1 - Gas Exist Fac Des & Inst").Select
    Range("D14").Select
    ActiveWorkbook.Save
    Range("D15").Select
    ActiveWindow.ActivateNext
    Range("D12").Select
    ActiveCell.Comment.Visible = True
    Range("D15").Select
    ActiveCell.Comment.Visible = False
    Range("D23").Select
    ActiveWorkbook.Save
    Range("D13").Select
    ActiveWorkbook.Save
    Range("D13").Select
    ActiveWorkbook.Save
    Range("D15").Select
    ActiveWorkbook.Save
    Range("D20").Select
    ActiveWorkbook.Save
    Range("D20").Select
    ActiveWorkbook.Save
    Range("D11").Select
    ActiveWorkbook.Save
    Range("D21").Select
End Sub


Sub test()

Dim filename As String
Dim fullRangeString As String

Dim returnValue As Variant
Dim wb As Workbook
Dim ws As Worksheet

Dim rng As Range

    'get workbook path
    filename = Application.GetOpenFilename(FileFilter:="Excel Files (*.xls), *.xls", Title:="Please select a file")


    'set our workbook and open it
    Set wb = Application.Workbooks.Open(filename)

    'set our worksheet
    Set ws = wb.Worksheets("Sheet1")

    'set the range for vlookup
    Set rng = ws.Range("A3:I13")


    'Do what you need to here with the range (will get error (unable to get vlookup property of worksheet) if value doesn't exist
    returnValue = Application.WorksheetFunction.VLookup("test4", rng, 2, False)


    MsgBox returnValue
    'If you need a fully declared range string for use in a vlookup formula, then
    'you'll need something like this (this won't work if there is any spaces or special
    'charactors in the sheet name



    'fullRangeString = "[" & rng.Parent.Parent.Name & "]" _
                        & rng.Parent.Name & "!" & rng.Address

    'Sheet1.Cells(10, 10).Formula = "=VLOOKUP(A1," & fullRangeString & ",8,False)"




    'close workbook if you need to
    wb.Close False


End Sub

