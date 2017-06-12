Attribute VB_Name = "Reset_Summaries"

Sub Reset_Findings()
Attribute Reset_Findings.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Reset_Findings Macro
'

'
    Range("T8:Z8").Select
    Application.CutCopyMode = False
    Selection.ListObject.ListColumns(13).Delete
    Selection.ListObject.ListColumns(13).Delete
    Selection.ListObject.ListColumns(13).Delete
    Selection.ListObject.ListColumns(13).Delete
    Selection.ListObject.ListColumns(13).Delete
    Selection.ListObject.ListColumns(13).Delete
    Selection.ListObject.ListColumns(13).Delete
    Range("NCESummary[[#Headers],[ABBT0045481]]").Select
    Application.GoTo Reference:="FacIDs"
    Selection.Copy
    Sheets("Findings Summary").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("R8").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Application.CutCopyMode = False
    Selection.FillRight
    ActiveWindow.ScrollRow = 157
    ActiveWindow.ScrollRow = 155
    ActiveWindow.ScrollRow = 152
    ActiveWindow.ScrollRow = 146
    ActiveWindow.ScrollRow = 142
    ActiveWindow.ScrollRow = 139
    ActiveWindow.ScrollRow = 136
    ActiveWindow.ScrollRow = 132
    ActiveWindow.ScrollRow = 130
    ActiveWindow.ScrollRow = 128
    ActiveWindow.ScrollRow = 126
    ActiveWindow.ScrollRow = 125
    ActiveWindow.ScrollRow = 123
    ActiveWindow.ScrollRow = 122
    ActiveWindow.ScrollRow = 120
    ActiveWindow.ScrollRow = 119
    ActiveWindow.ScrollRow = 117
    ActiveWindow.ScrollRow = 115
    ActiveWindow.ScrollRow = 114
    ActiveWindow.ScrollRow = 112
    ActiveWindow.ScrollRow = 110
    ActiveWindow.ScrollRow = 108
    ActiveWindow.ScrollRow = 106
    ActiveWindow.ScrollRow = 103
    ActiveWindow.ScrollRow = 102
    ActiveWindow.ScrollRow = 99
    ActiveWindow.ScrollRow = 97
    ActiveWindow.ScrollRow = 95
    ActiveWindow.ScrollRow = 93
    ActiveWindow.ScrollRow = 91
    ActiveWindow.ScrollRow = 88
    ActiveWindow.ScrollRow = 86
    ActiveWindow.ScrollRow = 84
    ActiveWindow.ScrollRow = 81
    ActiveWindow.ScrollRow = 79
    ActiveWindow.ScrollRow = 77
    ActiveWindow.ScrollRow = 74
    ActiveWindow.ScrollRow = 72
    ActiveWindow.ScrollRow = 70
    ActiveWindow.ScrollRow = 67
    ActiveWindow.ScrollRow = 66
    ActiveWindow.ScrollRow = 64
    ActiveWindow.ScrollRow = 62
    ActiveWindow.ScrollRow = 61
    ActiveWindow.ScrollRow = 59
    ActiveWindow.ScrollRow = 58
    ActiveWindow.ScrollRow = 56
    ActiveWindow.ScrollRow = 54
    ActiveWindow.ScrollRow = 53
    ActiveWindow.ScrollRow = 52
    ActiveWindow.ScrollRow = 51
    ActiveWindow.ScrollRow = 50
    ActiveWindow.ScrollRow = 48
    ActiveWindow.ScrollRow = 47
    ActiveWindow.ScrollRow = 45
    ActiveWindow.ScrollRow = 43
    ActiveWindow.ScrollRow = 42
    ActiveWindow.ScrollRow = 41
    ActiveWindow.ScrollRow = 40
    ActiveWindow.ScrollRow = 39
    ActiveWindow.ScrollRow = 38
    ActiveWindow.ScrollRow = 37
    ActiveWindow.ScrollRow = 36
    ActiveWindow.ScrollRow = 37
    ActiveWindow.ScrollRow = 39
    ActiveWindow.ScrollRow = 40
    ActiveWindow.ScrollRow = 41
    ActiveWindow.ScrollRow = 42
    ActiveWindow.ScrollRow = 43
    ActiveWindow.ScrollRow = 44
    ActiveWindow.ScrollRow = 46
    ActiveWindow.ScrollRow = 47
    ActiveWindow.ScrollRow = 49
    ActiveWindow.ScrollRow = 51
    ActiveWindow.ScrollRow = 52
    ActiveWindow.ScrollRow = 53
    ActiveWindow.ScrollRow = 54
    ActiveWindow.ScrollRow = 57
    ActiveWindow.ScrollRow = 58
    ActiveWindow.ScrollRow = 59
    ActiveWindow.ScrollRow = 57
    ActiveWindow.ScrollRow = 54
    ActiveWindow.ScrollRow = 51
    ActiveWindow.ScrollRow = 49
    ActiveWindow.ScrollRow = 46
    ActiveWindow.ScrollRow = 44
    ActiveWindow.ScrollRow = 42
    ActiveWindow.ScrollRow = 40
    ActiveWindow.ScrollRow = 37
    ActiveWindow.ScrollRow = 34
    ActiveWindow.ScrollRow = 30
    ActiveWindow.ScrollRow = 26
    ActiveWindow.ScrollRow = 23
    ActiveWindow.ScrollRow = 19
    ActiveWindow.ScrollRow = 13
    ActiveWindow.ScrollRow = 9
    ActiveWindow.ScrollRow = 6
    ActiveWindow.ScrollRow = 2
    ActiveWindow.ScrollRow = 1
End Sub
Sub Reset_Conclusion_Formula()
Attribute Reset_Conclusion_Formula.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Reset_Conclusion_Formula Macro
'

'
    Range("H11").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(COUNTIF(GasExFac[@[ABBT4120030]:[ABGS0002301]],""N/A"")=COLUMNS(GasExFac[@[ABBT4120030]:[ABGS0002301]]),""Not Applicable to all facilities in the property."","""")"
    Range("GasExFac[Reason for Conclusion]").FormulaR1C1 = _
        "=IF(COUNTIF(GasExFac[@[ABBT4120030]:[ABGS0002301]],""N/A"")=COLUMNS(GasExFac[@[ABBT4120030]:[ABGS0002301]]),""Not Applicable to all facilities in the property."","""")"
End Sub
