Attribute VB_Name = "ribbonFunctions"




Public Sub OpenVisio()
    Dim visioApp As Visio.Application
    Dim vsoDocument As Visio.Document
    
    Set visioApp = New Visio.Application
    visioApp.Visible = True

    Set vsoDocument = visioApp.Documents.Add("")
End Sub
