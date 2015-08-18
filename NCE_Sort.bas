Attribute VB_Name = "NCE_Sort"
Sub NCE_Sort()
Attribute NCE_Sort.VB_ProcData.VB_Invoke_Func = " \n14"
'
' NCE_Sort Macro
'

'
    ActiveWorkbook.Worksheets("NCE Component").ListObjects("NCESub").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("NCE Component").ListObjects("NCESub").Sort. _
        SortFields.Add Key:=Range("NCESub[Business Process]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("NCE Component").ListObjects("NCESub").Sort. _
        SortFields.Add Key:=Range("NCESub[Theme]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("NCE Component").ListObjects("NCESub").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

End Sub
