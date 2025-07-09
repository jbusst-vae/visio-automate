Attribute VB_Name = "shapes"





Public Sub AddCPoint(ByRef shp As Visio.Shape, ByVal x As Double, ByVal y As Double)
    Dim sec As Visio.Section
    Dim row As Integer

    ' Make sure shape has a connection point section
    If shp.SectionExists(visSectionConnectionPts, visExistsLocally) = 0 Then
        shp.AddSection visSectionConnectionPts
    End If

    ' Add a row for the new connection point
    row = shp.AddRow(visSectionConnectionPts, visRowLast, visTagCnnctPt)

    ' Set the X and Y cells (local coordinates)
    shp.CellsSRC(visSectionConnectionPts, row, visCnnctX).FormulaU = CStr(x)
    shp.CellsSRC(visSectionConnectionPts, row, visCnnctY).FormulaU = CStr(y)

    ' Optional: set connection point type (0 = inward/outward)
    shp.CellsSRC(visSectionConnectionPts, row, visCnnctDirX).FormulaU = "0"
    shp.CellsSRC(visSectionConnectionPts, row, visCnnctDirY).FormulaU = "0"

    Debug.Print "Added CPoint at " & CStr(x) & " " & CStr(y)
End Sub



' Function to create a custom module shape with specific pin labels
Public Function NewShape(ByRef page As Visio.Page, x As Double, y As Double, _
                           width As Double, height As Double, moduleName As String, _
                           Optional pinLabels As Variant) As Visio.Shape

    Dim moduleShape As Visio.Shape
    Set moduleShape = page.DrawRectangle(x, y, x + width, y + height)
    
    With moduleShape
        .Text = moduleName
        .CellsU("FillForegnd").Formula = "RGB(240,240,240)" ' Light gray fill
        .CellsU("LineColor").Formula = "RGB(0,0,0)" ' Black border
        .CellsU("LineWeight").Formula = "1.5 pt"
        .Name = moduleName
    End With

    ' draw pins, if we have any
    If Not IsMissing(pinLabels) and IsArray(pinLabels) Then
        Dim i As Integer
        Dim pinShape As Visio.Shape

        For i = 0 To UBound(pinLabels)
            ' Create small rectangles for pins along the right edge
            Set pinShape = page.DrawRectangle(x - 0.45, y - i*0.3, x - 0.2, y - 0.1 - i*0.3)
            With pinShape
                .Text = pinLabels(i)
                .CellsU("Char.Size").Result("pt") = 6
                .CellsU("FillForegnd").Formula = "RGB(255,255,255)"
                .CellsU("LineColor").Formula = "RGB(0,0,0)"
                .CellsU("LineWeight").Formula = "0.2 pt"
                .Name = moduleName & "_Pin" & (i + 1)
            End With
        Next i
    End If
    
    Set NewShape = moduleShape
End Function



Public Function DropCustomShape(ByRef app As Visio.Application, ByRef page as Visio.Page, shape_opt As Integer, x as integer, y as integer) As Visio.shape
    Dim stencil As Visio.Document
    Dim master As Visio.Master
    Dim shape_str as string
    Dim droppedShape as visio.shape
    
    shape_str = replace("MSSBN","N",cstr(shape_opt))

    Set stencil = app.Documents.OpenEx("C:\Users\jbusst\Documents - local\visiovba\MSSB-stencil.vssx", visOpenHidden)
    Set master = stencil.Masters(shape_str)
    Set droppedShape = page.Drop(master, x, y)

    ' Move it so its bottom-left corner is at (x, y)
    Dim w As Double, h As Double
    w = droppedShape.CellsU("Width").ResultIU
    h = droppedShape.CellsU("Height").ResultIU

    ' Set its Pin to bottom left
    droppedShape.CellsU("PinX").FormulaU = x + w / 2
    droppedShape.CellsU("PinY").FormulaU = y + h / 2

    set DropCustomShape = droppedShape
End Function

Public Sub MSSB1(ByRef app As Visio.Application, ByRef page as Visio.Page)
    Dim x(4) As Float
    Dim y(4) As Float

    x = Array(169.69, 169.69)
    y = Array(22.62, 29)
    
End Sub

Public Sub MSSB3(ByRef app As Visio.Application, ByRef page as Visio.Page)
    Dim x(4) As Float
    Dim y(4) As Float

    x = Array(5.9171, 10.3171, 150.8739, )
    y = Array(22.62, 29)
End Sub

Public Sub Connect1(shape1 As Visio.Shape, shape2 As Visio.Shape, direction As String, pin1 As Integer, pin2 As Integer)
    Dim connector As Visio.Shape
    Set connector = shape1.ContainingPage.DrawLine(0, 0, 1, 1)

    ' connect pins
    connector.Cells("BeginX").GlueTo shape1.Cells(Replace("Connections.XN", "N", CStr(pin1)))
    connector.Cells("BeginY").GlueTo shape1.Cells(Replace("Connections.YN", "N", CStr(pin1)))
    connector.Cells("EndX").GlueTo   shape2.Cells(Replace("Connections.XN", "N", CStr(pin2)))
    connector.Cells("EndY").GlueTo   shape2.Cells(Replace("Connections.YN", "N", CStr(pin2)))
    
    connector.Cells("ObjType").Formula = "2"
    connector.Cells("ConLineRouteExt").Formula = "1"    ' route with right angles
    connector.Cells("ShapeRouteStyle").Formula = visLORouteRightAngle       ' Route using only horizontal and vertical (ie right angles)
    ' connector.Cells("PlowCode").Formula = "0"         ' Plow through objects or not

    ' Set line jumps to show when lines cross
    connector.Cells("ConLineJumpStyle").Formula = "2"       ' 0=no jump, 1=arc jumps, 2=square jumps
    connector.Cells("ConLineJumpCode").Formula = "1"        ' 0=no jump, 1=always jump over, 2=always jump under, 3=auto
    
    ' Control routing behavior
    connector.Cells("ConFixedCode").Formula = "0"           ' 0=reroute when shape moved, 1=fixed connector, 2=partically fixed
    connector.Cells("ConLineJumpDirX").Formula = "0"        ' 0=auto, 1=force jump in positive x, -1=force jumps in negative x
    connector.Cells("ConLineJumpDirY").Formula = "0"        ' 0=auto, 1=force jump in positive y, -1=force jumps in negative y
    
    
    ' Style the line
    connector.Cells("LinePattern").Formula = "2"
    connector.Cells("LineColor").Formula = "2"
End Sub

Public Sub Connect2(ByRef shape1 As Visio.Shape, ByRef shape2 As Visio.Shape, direction As String, pinNo As Integer)
    Dim connector As Visio.Shape
    Set connector = shape1.ContainingPage.DrawLine(0, 0, 1, 1)
    
    ' Connect to your custom connection points
    connector.Cells("BeginX").GlueTo shape1.Cells(Replace("Connections.XN", "N", CStr(pinNo)))
    connector.Cells("BeginY").GlueTo shape1.Cells(Replace("Connections.YN", "N", CStr(pinNo)))
    connector.Cells("EndX").GlueTo   shape2.Cells(Replace("Connections.XN", "N", CStr(pinNo)))
    connector.Cells("EndY").GlueTo   shape2.Cells(Replace("Connections.YN", "N", CStr(pinNo)))
    
    ' Set the line to use right-angle routing
    connector.Cells("ConLineRouteExt").Formula = "1"  ' Enable right-angle routing
    
    ' Force the initial direction from the connection point
    ' Select Case UCase(direction)
    '     Case "right"
    '         connector.Cells("ConLineJumpDirX").Formula = "1"   ' Force right initially
    '         connector.Cells("BeginTrigger").Formula = "0"      ' Horizontal first
    '     Case "left"
    '         connector.Cells("ConLineJumpDirX").Formula = "-1"  ' Force left initially
    '         connector.Cells("BeginTrigger").Formula = "0"      ' Horizontal first
    '     Case "up"
    '         connector.Cells("ConLineJumpDirY").Formula = "1"   ' Force up initially
    '         connector.Cells("BeginTrigger").Formula = "1"      ' Vertical first
    '     Case "down"
    '         connector.Cells("ConLineJumpDirY").Formula = "-1"  ' Force down initially
    '         connector.Cells("BeginTrigger").Formula = "1"      ' Vertical first
    ' End Select
    
    connector.Cells("ShapeRouteStyle").Formula = "16" ' Connector style
    connector.Cells("ConLineJumpStyle").Formula = "1"  ' Arc jumps
    
    connector.Cells("LinePattern").Formula = "2"
    connector.Cells("LineColor").Formula = "2"
End Sub

Public Sub Connect3(ByRef shape1 As Visio.Shape, ByRef shape2 As Visio.Shape, direction As String, pinNumber As Integer)
    Dim connector As Shape
    Dim initialDirection As Integer
    
    ' Set initial direction based on input parameter
    ' Select Case LCase(direction)
    '     Case "left"
    '         initialDirection = msoConnectorStraight ' Will be modified by routing
    '     Case "right"
    '         initialDirection = msoConnectorStraight
    '     Case "up"
    '         initialDirection = msoConnectorStraight
    '     Case "down"
    '         initialDirection = msoConnectorStraight
    '     Case Else
    '         initialDirection = msoConnectorStraight
    ' End Select

    initialDirection = msoConnectorStraight

    ' Create the connector
    Set connector = shape1.Parent.Shapes.AddConnector(msoConnectorStraight, 0, 0, 100, 100)
    
    ' Connect the shapes using specified pin numbers
    connector.ConnectorFormat.BeginConnect shape1, pinNumber
    connector.ConnectorFormat.EndConnect shape2, pinNumber
    
    ' Set connector properties
    With connector.Line
        .DashStyle = msoLineDash
        .ForeColor.RGB = RGB(255, 0, 0) ' Red color
        .Weight = 1.5
    End With
    
    ' Set routing properties
    With connector.ConnectorFormat
        .Type = msoConnectorElbow ' Right angle routing
    End With
    
    ' Enable smart routing and arc jumps
    connector.ConnectorFormat.Type = msoConnectorElbow
    
    ' Set initial direction by adjusting connector type based on direction parameter
    Select Case LCase(direction)
        Case "left"
            connector.ConnectorFormat.Type = msoConnectorElbow
        Case "right"
            connector.ConnectorFormat.Type = msoConnectorElbow
        Case "up"
            connector.ConnectorFormat.Type = msoConnectorElbow
        Case "down"
            connector.ConnectorFormat.Type = msoConnectorElbow
    End Select
    
    ' Return the connector shape
    ' Set ConnectShapes = connector
End Sub


