Attribute VB_Name = "main"



Public Type VApp
    app  As Visio.Application
    doc  As Visio.Document
    page As Visio.Page
End Type


Public Sub Main()
    Dim vis As VApp
    vis = GetVisio()

    ' ExcelDiagram vis.app, vis.page

    Dim shape1 As Visio.Shape, shape2 As Visio.Shape
    Set shape1 = NewShape(vis.page, 0, 1, 1, 1, "shape 1")

    For i = 0 To 4
        AddCPoint shape1, 1, i*0.2
    Next i

    Set shape2 = NewShape(vis.page, 5, 4, 1, 1, "shape 2")

    For i = 0 To 4
        AddCPoint shape2, 0, i*0.2
    Next i
    
    Connect1 shape1, shape2, "right", 1, 1
    Connect1 shape1, shape2, "right", 1, 2
    Connect1 shape1, shape2, "right", 3, 3
    Connect1 shape1, shape2, "right", 4, 4
End Sub

Sub SampleDiagram()
    ' Declare Visio application and document objects
    Dim visApp As Visio.Application
    Dim visDoc As Visio.Document
    Dim visPage As Visio.Page
    Dim visShape1 As Visio.Shape
    Dim visShape2 As Visio.Shape
    Dim visConnector As Visio.Shape
    
    ' Error handling
    On Error GoTo ErrorHandler
    
    ' Create or get Visio application
    On Error Resume Next
    Set visApp = GetObject(, "Visio.Application")
    If visApp Is Nothing Then
        Set visApp = CreateObject("Visio.Application")
    End If
    On Error GoTo ErrorHandler
    
    ' Make Visio visible
    visApp.Visible = True
    
    ' Create a new blank document
    Set visDoc = visApp.Documents.Add("")
    Set visPage = visDoc.Pages(1)
    
    ' Set page properties (optional)
    visPage.Name = "Wiring Diagram"
    
    ' Create first block (rectangle)
    ' Parameters: Left, Bottom, Right, Top (in inches)
    Set visShape1 = visPage.DrawRectangle(1, 6, 3, 7)
    With visShape1
        .Text = "Module 1" & vbCrLf & "Pin A1"
        .CellsU("FillForegnd").Formula = "RGB(200,220,255)" ' Light blue fill
        .CellsU("LineColor").Formula = "RGB(0,0,0)" ' Black border
        .CellsU("LineWeight").Formula = "2 pt" ' Border thickness
        .Name = "Module1"
    End With
    
    ' Create second block (rectangle)
    Set visShape2 = visPage.DrawRectangle(6, 6, 8, 7)
    With visShape2
        .Text = "Module 2" & vbCrLf & "Pin B1"
        .CellsU("FillForegnd").Formula = "RGB(255,220,200)" ' Light orange fill
        .CellsU("LineColor").Formula = "RGB(0,0,0)" ' Black border
        .CellsU("LineWeight").Formula = "2 pt" ' Border thickness
        .Name = "Module2"
    End With
    
    ' Create connector line between the blocks
    Set visConnector = visPage.DrawLine(3, 6.5, 6, 6.5)
    With visConnector
        .CellsU("LineColor").Formula = "RGB(255,0,0)" ' Red line
        .CellsU("LineWeight").Formula = "1.5 pt" ' Line thickness
        .Name = "Connection1"
        ' Add arrowhead at the end
        .CellsU("EndArrow").Formula = "5" ' Standard arrowhead
    End With
    
    ' Add connection label
    Dim visLabel As Visio.Shape
    Set visLabel = visPage.DrawRectangle(4, 6.8, 5, 7.2)
    With visLabel
        .Text = "Wire 1"
        .CellsU("FillForegnd").Formula = "RGB(255,255,255)" ' White fill
        .CellsU("LineColor").Formula = "RGB(100,100,100)" ' Gray border
        .CellsU("LineWeight").Formula = "0.5 pt" ' Thin border
        .Name = "WireLabel1"
    End With
    
    ' Auto-fit page to content
    visPage.ResizeToFitContents
    
    ' Zoom to fit the page
    visApp.ActiveWindow.ViewFit = visFitPage
    
    MsgBox "Basic wiring diagram created successfully!", vbInformation, "Visio VBA"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Visio VBA Error"
    
End Sub

' Example of how to read data from Excel and create diagram
Sub ExcelDiagram(visApp As Visio.Application, page As Visio.Page)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Double
    Dim shape As Visio.Shape
    
    ' Set reference to your Excel worksheet
    Set ws = ThisWorkbook.Worksheets("WiringData") ' Change to your sheet name
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To lastRow ' Assuming row 1 has headers
        Dim moduleName As String
        
        moduleName  = CStr(ws.Cells(i, 1).Value)
        inputPins   = Split(CStr(ws.Cells(i, 2).Value), ", ")
        outputPins  = Split(CStr(ws.Cells(i, 3).Value), ", ")
        connections = Split(CStr(ws.Cells(i, 4).Value), ", ")

        Set shape = NewShape(page, i*1.5, i*1.5, 1, 1, moduleName, inputPins)
    Next i

    ' Auto-fit page to content
    page.ResizeToFitContents
    
    ' Zoom to fit the page
    visApp.ActiveWindow.ViewFit = visFitPage
End Sub



' grabs active visio app and page
Public  Function GetVisio() As VApp
    Dim visApp As Visio.Application
    Dim visDoc As Visio.Document
    Dim visPage As Visio.Page
    Dim visShape1 As Visio.Shape
    Dim visShape2 As Visio.Shape
    Dim visConnector As Visio.Shape
    
    ' Create or get Visio application
    On Error Resume Next
    Set visApp = GetObject(, "Visio.Application")
    If visApp Is Nothing Then
        Set visApp = CreateObject("Visio.Application")
        Set visDoc = visApp.Documents.Add("")
        Set visPage = visDoc.Pages(1)
        GoTo PassBack
    End If
    On Error GoTo 0
    
     ' Get the active document
    If visApp.Documents.Count = 0 Then
        MsgBox "No open Visio document found. Please open one.", vbExclamation
        Exit Function
    End If
    Set visDoc = visApp.ActiveDocument

    ' Get the active page
    If visDoc.Pages.Count = 0 Then
        MsgBox "No pages found in the active document.", vbExclamation
        Exit Function
    End If
    Set visPage = visApp.ActivePage

    PassBack:
    visApp.Visible = True
    ClearPage visPage

    Set GetVisio.app = visApp
    Set GetVisio.doc = visDoc
    Set GetVisio.page = visPage
End Function
