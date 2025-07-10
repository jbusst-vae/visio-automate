Attribute VB_Name = "main"



Public Type VApp
    app  As Visio.Application
    doc  As Visio.Document
    page As Visio.Page
End Type


Public Sub Main()
    Dim vis As VApp
    vis = GetVisio()

    Dim shape1 as visio.shape
    Dim shape2 as visio.shape
    Dim shape3 as visio.shape

    set shape1 = DropCustomShape(vis.app, vis.page, 1, 13, 3)
    set shape2 = DropCustomShape(vis.app, vis.page, 3, 2, 3)
    ' set shape3 = DropCustomShape(vis.app, vis.page, 3, 24, 0)

    DrawLineFromShape shape1
    DrawLineFromShape shape2

    dim connector as visio.shape
    set connector = connect1(shape1, shape2, "right", 1, 1)

    Debug.print "\n- Start application -"

    dim x As Double
    dim y As Double
    dim index As Integer
    
    index = 0

    for i = 1 to 100
        x = connector.CellsSRC(10, i, 0).Result("mm")
        y = connector.CellsSRC(10, i, 1).Result("mm")

        if x = 0 and y = 0 and i > 1 then
            index = i
            exit for
        end if 

        ' debug.print x
        ' debug.print y

    next i

    ' ' shuffle end point down one space
    ' connector.cellssrc(10, index, 0).formulau = connector.cellssrc(10, index-1, 0).formulau
    ' connector.cellssrc(10, index, 1).formulau = connector.cellssrc(10, index-1, 1).formulau

    ' ' add a new points in
    ' connector.cellssrc(10, index-2, 0).formulau = "-500"
    ' connector.cellssrc(10, index-2, 1).formulau = "50"
    connector.addrow 10, index, visTagLineTo 
    connector.addrow 10, index+1, visTagLineTo 
    connector.addrow 10, index+1, visTagLineTo 
    connector.cellssrc(10, index, 0).formulau = "-500"
    connector.cellssrc(10, index, 1).formulau = "50"

    for i = 1 to 100
        x = connector.CellsSRC(10, i, visX).Result("mm")
        y = connector.CellsSRC(10, i, visY).Result("mm")

        if x = 0 and y = 0 and i > 1 then
            exit for
        end if 

        debug.print x
        debug.print y

    next i


    ' Debug.print connector.CellsSRC(visSectionFirst + 1, 1, 0).Result("mm")
    ' Debug.print connector.CellsSRC(visSectionFirst + 1, 1, 1).Result("mm")
    

    ' ' set cell = connector.CellsSRC(visSectionFirst + 1, 1, 1)
    ' ' connector.CellsSRC(visSectionFirst + 1, 1, 1).Result("mm").formulau = "1 in"
    
    ' Debug.print connector.CellsSRC(visSectionFirst + 1, 2, 0).Result("mm")
    ' Debug.print connector.CellsSRC(visSectionFirst + 1, 2, 1).Result("mm")

    ' Debug.print connector.CellsSRC(visSectionFirst + 1, 3, 0).Result("mm")
    ' Debug.print connector.CellsSRC(visSectionFirst + 1, 3, 1).Result("mm")

    
    ' for i = 0 To shape1.RowCount(visSectionConnectionPts) - 1
    '     Connect1 shape1, shape2, "right", i+1, 1
    ' Next i

    ' connect1 shape2, shape3, "right", 1, 1
    ' connect1 shape2, shape3, "right", 3, 2
    ' connect1 shape2, shape3, "right", 2, 5
    

    ' Dim shape1 As Visio.Shape, shape2 As Visio.Shape
    ' Set shape1 = NewShape(vis.page, 0, 1, 1, 1, "shape 1")

    ' For i = 0 To 4
    '     AddCPoint shape1, 1, i*0.2
    ' Next i

    ' Set shape2 = NewShape(vis.page, 5, 4, 1, 1, "shape 2")

    ' For i = 0 To 4
    '     AddCPoint shape2, 0, i*0.2
    ' Next i
    
    ' Connect1 shape1, shape2, "right", 1, 1
    ' Connect1 shape1, shape2, "right", 1, 2
    ' Connect1 shape1, shape2, "right", 3, 3
    ' Connect1 shape1, shape2, "right", 4, 4
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



Sub DrawLineFromShape(shape As Visio.Shape)
    Dim page As Visio.Page
    Set page = shape.Parent
    
    ' Get the local coordinates of the first connection point
    Dim connX As Double
    Dim connY As Double
    connX = shape.Cells("Connections.X1")
    connY = shape.Cells("Connections.Y1")
    
    ' Calculate the page coordinates of the connection point
    Dim startX As Double
    Dim startY As Double
    startX = shape.Cells("PinX") + connX - shape.Cells("LocPinX")
    startY = shape.Cells("PinY") + connY - shape.Cells("LocPinY")
    
    ' Define the points for the line path
    Dim points(1 To 10) As Double
    points(1) = startX        ' Starting point
    points(2) = startY
    points(3) = startX + 2    ' Extend right by 2 units
    points(4) = startY
    points(5) = startX + 2    ' Down by 1 unit
    points(6) = startY - 1
    points(7) = startX + 3    ' Right by 1 unit
    points(8) = startY - 1
    points(9) = startX + 3    ' Down by 1 unit
    points(10) = startY - 2
    
    ' ' Draw the polyline on the page
    ' Dim line As Visio.Shape
    ' Set line = shape.containingpage.DrawPolyline(points, 0)
    
    ' ' Glue the beginning of the line to the shape's first connection point
    ' line.Cells("BeginX").GlueTo shape.Cells("Connections.X1")
    ' line.Cells("BeginY").GlueTo shape.Cells("Connections.Y1")
End Sub