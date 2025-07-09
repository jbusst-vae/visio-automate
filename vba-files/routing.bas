attribute vb_name = "routing"



Sub RouteConnectorsToMinimizeOverlaps()
    Dim page As Visio.Page
    Dim visApp As Visio.Application
    Dim shape As Visio.shape
    Dim connector As Visio.shape
    Dim gridSize As Double
    Dim obstacles As Collection
    Dim startPoint As Variant, endPoint As Variant
    Dim path As Collection
    
    Set visApp = GetObject(, "Visio.Application")
    Set page = visApp.ActivePage
    Set obstacles = New Collection
    gridSize = 0.1 ' Grid cell size in inches
    
    ' Step 1: Collect obstacles (shapes and existing connectors)
    For Each shape In page.Shapes
        If shape.OneD = 0 Then ' 2D shape (not a connector)
            AddShapeToObstacles shape, obstacles, gridSize
        End If
    Next shape
    
    ' Step 2: Route each connector
    For Each connector In page.Shapes
        If connector.OneD = 1 Then ' 1D shape (connector)
            ' Get start and end points from connected shapes
            startPoint = GetConnectionPoint(connector, "BeginX", "BeginY")
            endPoint = GetConnectionPoint(connector, "EndX", "EndY")
            
            ' Find path avoiding obstacles
            Set path = FindPath(startPoint, endPoint, obstacles, gridSize)
            
            ' Apply path to connector
            SetConnectorPath connector, path
            
            ' Add this connector to obstacles for the next iteration
            AddConnectorToObstacles connector, obstacles, gridSize
        End If
    Next connector
End Sub

' Helper: Add shape to obstacle grid
Sub AddShapeToObstacles(shape As Visio.shape, obstacles As Collection, gridSize As Double)
    Dim x As Double, y As Double
    Dim width As Double, height As Double
    x = shape.Cells("PinX").ResultIU - shape.Cells("Width").ResultIU / 2
    y = shape.Cells("PinY").ResultIU - shape.Cells("Height").ResultIU / 2
    width = shape.Cells("Width").ResultIU
    height = shape.Cells("Height").ResultIU
    
    Dim i As Integer, j As Integer
    For i = 0 To CInt(width / gridSize)
        For j = 0 To CInt(height / gridSize)
            obstacles.Add Array(CInt(x / gridSize) + i, CInt(y / gridSize) + j)
        Next j
    Next i
End Sub

' Helper: Get connection point coordinates
Function GetConnectionPoint(connector As Visio.shape, xCell As String, yCell As String) As Variant
    Dim x As Double, y As Double
    x = connector.Cells(xCell).ResultIU
    y = connector.Cells(yCell).ResultIU
    GetConnectionPoint = Array(x, y)
End Function

' Helper: Add connector to obstacles (simplified)
Sub AddConnectorToObstacles(connector As Visio.shape, obstacles As Collection, gridSize As Double)
    Dim points As Variant
    Dim i As Integer
    points = connector.PathPoints(0.01) ' Get points along path
    For i = 0 To UBound(points, 1)
        obstacles.Add Array(CInt(points(i, 0) / gridSize), CInt(points(i, 1) / gridSize))
    Next i
End Sub

' Placeholder: Find path using A* (simplified)
Function FindPath(startPoint As Variant, endPoint As Variant, obstacles As Collection, gridSize As Double) As Collection
    ' Implement A* here or use a library function
    ' ForOsmatic A* returns a list of (x, y) coordinates
    Set FindPath = New Collection
    ' Add dummy path for demonstration
    FindPath.Add startPoint
    ' Add intermediate points avoiding obstacles (requires A* logic)
    FindPath.Add Array(startPoint(0) + 0.5, startPoint(1) + 0.5)
    GetAStarPath startPoint, endPoint, obstacles, gridSize, FindPath
    FindPath.Add endPoint
End Function

' Placeholder: Apply path to connector
Sub SetConnectorPath(connector As Visio.shape, path As Collection)
    Dim i As Integer
    ' Clear existing geometry
    If connector.RowExists(visSectionFirstComponent, 1) Then
        connector.DeleteRow visSectionFirstComponent, 1
    End If
    
    ' Add new geometry section
    connector.AddSection visSectionFirstComponent + 1
    connector.AddRow visSectionFirstComponent + 1, visRowComponent, visTagComponent
    
    ' Add path points
    For i = 0 To path.Count - 1
        connector.AddRow visSectionFirstComponent + 1, visRowVertex, IIf(i = 0, visTagMoveTo, visTagLineTo)
        connector.CellsSRC(visSectionFirstComponent + 1, i, visX).FormulaU = path(i)(0)
        connector.CellsSRC(visSectionFirstComponent + 1, i, visY).FormulaU = path(i)(1)
    Next i
End Sub

' Placeholder: A* pathfinding (simplified, needs full implementation)
Sub GetAStarPath(startPoint, endPoint, obstacles, gridSize, path As Collection)
    ' Add A* logic here to populate path with points avoiding obstacles
    ' For now, adds a straight path (replace with actual A*)
    Dim midX As Double, midY As Double
    midX = (startPoint(0) + endPoint(0)) / 2
    midY = (startPoint(1) + endPoint(1)) / 2
    path.Add Array(midX, midY)
End Sub