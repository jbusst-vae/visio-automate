Attribute VB_Name = "visio"



Sub ClearVisio(visDoc As Visio.Document)
    Dim visPage As Visio.Page

    ' Delete all pages except one
    Do While visDoc.Pages.Count > 1
        visDoc.Pages(visDoc.Pages.Count).Delete
    Loop

    ' Clear the shapes on the remaining page
    Set visPage = visDoc.Pages(1)

    ClearPage visPage
End Sub

Public Sub ClearPage(page As Visio.Page)
    Dim i As Integer

    For i = page.Shapes.Count To 1 Step -1
        page.Shapes(i).Delete
    Next i
End Sub

Public Sub OpenVisio()
    Dim visioApp As Visio.Application
    Dim vsoDocument As Visio.Document
    
    Set visioApp = New Visio.Application
    visioApp.Visible = True

    Set vsoDocument = visioApp.Documents.Add("")
End Sub


