Attribute VB_Name = "LogoChangeCode"
Sub LogoUpdate()
Attribute LogoUpdate.VB_ProcData.VB_Invoke_Func = " \n14"
'
' LogoUpdate Macro
'
    Dim Client As String, clLogo As Object, thesheets As Object
    
    ' make sure the worksheet activate event doesn't do it's whole thing
    Rebuild = True
    
    Client = "Statoil"
'
    Worksheets("Lookups").Activate
    Worksheets("Lookups").Shapes.Range(Array(Client)).Select
    Selection.Copy
    
    Set thesheets = ActiveWorkbook.Worksheets(Array("Handout", "Facility List"))
    For Each sht In thesheets
        sht.Activate
        sht.Shapes("thisWorkbookLogo").Delete
        sht.Range("a1").Select
        sht.Paste
        sht.Shapes(Client).Name = "thisWorkbookLogo"
    Next sht

    Set thesheets = ActiveWorkbook.Worksheets(Array("BP1 - Gas Exist Fac Des & Inst", _
        "BP2 - Gas New Fac Des & Inst", "BP3 - Gas Measurement", "BP4 - Gas Recording", _
        "BP5 - Gas Reporting", "BP6 - HC Liq Ex Fac Des & Inst", _
        "BP7 - HC Liq New Fac Des & Inst", "BP8 - HC Liquid Measurement", _
        "BP9 - HC Liquid Recording", "BP10 - HC Liquid Reporting", _
        "BP11 - Water Ex Fac Des & Inst", "BP12 - Water New Fac Des & Inst", _
        "BP13 - Water Measurement", "BP14 - Water Recording", "BP15 - Water Reporting"))

    
    For Each sht In thesheets
        sht.Activate
        sht.Shapes("thisWorkbookLogo").Delete
        sht.Range("c2").Select
        sht.Paste
        sht.Shapes(Client).Name = "thisWorkbookLogo"
    Next sht
    
    ' done, re-enable the worksheet activate event
    Rebuild = False
    
    
End Sub
Sub logoSelect()
Attribute logoSelect.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim logos As Object
    ' make sure the worksheet activate event doesn't do it's whole thing
    Rebuild = True
    
    Worksheets("Lookups").Activate
    Set logos = Worksheets("Lookups").Shapes
    
    For Each shp In logos
        Debug.Print shp.Name
    Next shp

    ' done, re-enable the worksheet activate event
    Rebuild = False


    
End Sub
