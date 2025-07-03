Attribute VB_Name = "claude"


Public  Sub Main()
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
    End If
    
    ' Make Visio visible
    visApp.Visible = True
    
    ' Create a new blank document
    Set visDoc = visApp.Documents.Add("")
    Set visPage = visDoc.Pages(1)
    
    ' Set page properties (optional)
    visPage.Name = "Wiring Diagram"

    CreateDiagramFromExcelData visApp, visPage
End Sub

Sub CreateBasicWiringDiagram()
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

Function AddConnectionPoint(shape As Visio.Shape, x As Double, y As Double) As Integer
    ' Add a connection point at specified coordinates relative to shape
    Dim connectionRow As Integer
    connectionRow = shape.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
    
    ' Set the X and Y coordinates for the connection point
    shape.CellsSRC(visSectionConnectionPts, connectionRow, visCnnctX).Formula = x
    shape.CellsSRC(visSectionConnectionPts, connectionRow, visCnnctY).Formula = y
    
    AddConnectionPoint = connectionRow
End Function

' Function to create a custom module shape with specific pin labels
Function CreateModuleShape(ByRef page As Visio.Page, left As Double, bottom As Double, _
                          width As Double, height As Double, moduleName As String, _
                          pinLabels As Variant) As Visio.Shape

    Dim moduleShape As Visio.Shape
    Set moduleShape = page.DrawRectangle(left, bottom, left + width, bottom + height)
    
    With moduleShape
        .Text = moduleName
        .CellsU("FillForegnd").Formula = "RGB(240,240,240)" ' Light gray fill
        .CellsU("LineColor").Formula = "RGB(0,0,0)" ' Black border
        .CellsU("LineWeight").Formula = "1.5 pt"
        .Name = moduleName
    End With

    ' Add pin labels (this is a simplified version - you'd expand based on your needs)
    If IsArray(pinLabels) Then
        Dim i As Integer
        Dim pinShape As Visio.Shape
        For i = 0 To UBound(pinLabels)
            ' Create small rectangles for pins along the right edge
            Set pinShape = page.DrawRectangle(left + width, bottom + (i * 0.3), _
                                            left + width + 0.5, bottom + (i * 0.3) + 0.2)
            With pinShape
                .Text = pinLabels(i)
                .CellsU("FillForegnd").Formula = "RGB(255,255,255)"
                .CellsU("LineColor").Formula = "RGB(0,0,0)"
                .CellsU("LineWeight").Formula = "0.5 pt"
                .Name = moduleName & "_Pin" & (i + 1)
            End With
        Next i
    End If
    
    Set CreateModuleShape = moduleShape
End Function

' Example of how to read data from Excel and create diagram
Sub CreateDiagramFromExcelData(visApp As Visio.Application, page As Visio.Page)
    ' This is a template for reading Excel data
    ' Assumes you have a worksheet with columns: Module1, Pin1, Module2, Pin2, WireLabel
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim shape As Visio.Shape
    
    ' Set reference to your Excel worksheet
    Set ws = ThisWorkbook.Worksheets("WiringData") ' Change to your sheet name
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To lastRow ' Assuming row 1 has headers
        Dim module1Name As String, pin1Name As String
        Dim module2Name As String, pin2Name As String
        Dim wireLabel As String
        
        module1Name = ws.Cells(i, 1).Value ' Column A
        pin1Name = ws.Cells(i, 2).Value    ' Column B
        module2Name = ws.Cells(i, 3).Value ' Column C
        pin2Name = ws.Cells(i, 4).Value    ' Column D
        wireLabel = ws.Cells(i, 5).Value   ' Column E

        Debug.Print module1Name
        
        ' Here you would create/find the modules and connect them
        ' This is where you'd implement your specific logic
        Dim pinLabels As Variant
        pinLabels = Array(pin1Name, pin2Name)


        Set shape = CreateModuleShape(page, i*10, i*10, 5, 5, module1Name, pinLabels)
    Next i

    ' Auto-fit page to content
    page.ResizeToFitContents
    
    ' Zoom to fit the page
    visApp.ActiveWindow.ViewFit = visFitPage

    MsgBox "Wiring diagram successfully created!",,"title" 
End Sub