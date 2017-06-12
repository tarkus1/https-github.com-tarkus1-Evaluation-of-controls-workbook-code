Attribute VB_Name = "ParamountSorts"
Sub SortForPmntCntls()
Attribute SortForPmntCntls.VB_ProcData.VB_Invoke_Func = " \n14"
'
' SortForPmntCntls Macro
'

'
    Rebuild = True
    
    Set oListObj = ActiveWorkbook.ActiveSheet.ListObjects(1)
    
    oListObj.Sort.SortFields.Clear
    
    oListObj.Sort. _
        SortFields.Add Key:=oListObj.ListColumns("NCE Component").DataBodyRange, SortOn:=xlSortOnValues _
        , Order:=xlAscending, DataOption:=xlSortNormal
    oListObj.Sort. _
        SortFields.Add Key:=oListObj.ListColumns("NCE").DataBodyRange, SortOn:=xlSortOnValues, Order _
        :=xlAscending, DataOption:=xlSortNormal
    With oListObj.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Rebuild = False
End Sub

Sub SortBackFromPmntCntls()
'
'
'
    Rebuild = True
    
    Set oListObj = ActiveWorkbook.ActiveSheet.ListObjects(1)
    
    oListObj.Sort.SortFields.Clear
    
    oListObj.Sort. _
        SortFields.Add Key:=oListObj.ListColumns("Theme").DataBodyRange, SortOn:=xlSortOnValues, Order _
        :=xlAscending, DataOption:=xlSortNormal
    oListObj.Sort. _
        SortFields.Add Key:=oListObj.ListColumns("NCE").DataBodyRange, SortOn:=xlSortOnValues, Order _
        :=xlAscending, DataOption:=xlSortNormal
    oListObj.Sort. _
        SortFields.Add Key:=oListObj.ListColumns("NCE Component").DataBodyRange, SortOn:=xlSortOnValues _
        , Order:=xlAscending, DataOption:=xlSortNormal
    With oListObj.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Rebuild = False
End Sub


