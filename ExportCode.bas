Attribute VB_Name = "ExportCode"
Sub xPort()
     
     ' export to C:\Users\Mark\OneDrive\MakeWaves\Corvelle\EPAP\Methodology\Code for version control with GIT
     
    Dim objMyProj As VBProject
    Dim objVBComp As VBComponent
     
    Set objMyProj = Application.VBE.ActiveVBProject
     
    For Each objVBComp In objMyProj.VBComponents
        If objVBComp.Type = vbext_ct_StdModule Then
            Debug.Print "module "; objVBComp.Name
            
            objVBComp.Export "C:\Users\Mark\OneDrive\MakeWaves\Corvelle\EPAP\Methodology\Evaluation of controls workbook code\" & objVBComp.Name & ".bas"
        End If
    Next
     
End Sub

