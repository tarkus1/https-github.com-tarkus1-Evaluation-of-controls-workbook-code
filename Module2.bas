Attribute VB_Name = "Module2"
Sub ncelist()
Attribute ncelist.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ncelist Macro
'

'
    Range("T1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$T$1:$Y$12"), , xlYes).Name = _
        "Table27"
    Range("W2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Application.CutCopyMode = False
    Range("U2").Select
End Sub
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'
    
    
    Range(Range("CC_Criteria").Offset(1, 0), Range("CC_Criteria").End(xlDown)).Clear
    
End Sub
Sub Macro3()
Attribute Macro3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro3 Macro
'

'
    Range("BP_NCEs[[#Headers],[NCEProd]]").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
End Sub
Sub clientcontrolfilter()
Attribute clientcontrolfilter.VB_ProcData.VB_Invoke_Func = " \n14"
'
' clientcontrolfilter Macro
'

'
    Range("ClientControls[[#Headers],[Theme]]").Select
    Range("ClientControls[#All]").AdvancedFilter Action:=xlFilterCopy, _
        CriteriaRange:=Range("'Client Controls'!Criteria"), CopyToRange:=Range( _
        "P1:S1"), Unique:=True
End Sub
