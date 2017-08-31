Attribute VB_Name = "ClientControlSorts"
Sub SortForClientCntls()
Attribute SortForClientCntls.VB_ProcData.VB_Invoke_Func = " \n14"
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
    Debug.Print ActiveSheet.Name; " paramount sort"
    
    Rebuild = False
End Sub

Sub SortBackFromClientCntls()
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
    Debug.Print ActiveSheet.Name; " back to normal sort"

    Rebuild = False
End Sub


