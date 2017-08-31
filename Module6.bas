Attribute VB_Name = "Module6"
Sub printPDF()
Attribute printPDF.VB_ProcData.VB_Invoke_Func = " \n14"
'
' printPDF Macro
'

'
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False, ActivePrinter:=PrintToPDF, _
        PrToFileName:="W:\NuVista\25 Evaluation of Controls\2017\sdf.pdf"

    
End Sub
Sub printpdf2()
Attribute printpdf2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' printpdf2 Macro
'

'
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
End Sub
