Attribute VB_Name = "ExportCode"
Sub xPort()
     
     ' export to C:\Users\Mark\OneDrive\MakeWaves\Corvelle\EPAP\Methodology\Code for version control with GIT
     
    Dim objMyProj As VBProject
    Dim objVBComp As VBComponent
    Dim fname As Variant
     
    Set objMyProj = Application.VBE.ActiveVBProject
     
    For Each objVBComp In objMyProj.VBComponents
        If objVBComp.Type = vbext_ct_StdModule Or objVBComp.Type = vbext_ct_ClassModule Then
            
            'fname = "C:\Users\Mark\OneDrive\MakeWaves\Corvelle\EPAP\Methodology\Evaluation of controls workbook code\" & objVBComp.Name & ".bas"
            fname = "C:\Users\mark_\OneDrive\MakeWaves\Corvelle\EPAP\Methodology\Evaluation of controls workbook code\" & objVBComp.Name & ".bas"
            Debug.Print "module "; fname
                                   
            objVBComp.Export (fname)
             
        End If
    Next
     
End Sub


