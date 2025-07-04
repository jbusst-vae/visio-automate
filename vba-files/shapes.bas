Attribute VB_Name = "shapes"





Sub AddCPoint(ByRef shp As Visio.Shape, ByVal x As Double, ByVal y As Double)
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
End Sub



' Function to create a custom module shape with specific pin labels
Function NewShape(ByRef page As Visio.Page, x As Double, y As Double, _
                           width As Double, height As Double, moduleName As String, _
                           pinLabels As Variant) As Visio.Shape

    Dim moduleShape As Visio.Shape
    Set moduleShape = page.DrawRectangle(x, y, x + width, y - height)
    
    With moduleShape
        .Text = moduleName
        .CellsU("FillForegnd").Formula = "RGB(240,240,240)" ' Light gray fill
        .CellsU("LineColor").Formula = "RGB(0,0,0)" ' Black border
        .CellsU("LineWeight").Formula = "1.5 pt"
        .Name = moduleName
    End With

    ' draw pins
    If IsArray(pinLabels) Then
        Dim i As Integer
        Dim pinShape As Visio.Shape

        For i = 0 To UBound(pinLabels)
            AddCPoint moduleShape, x, y - 0.15
            Debug.Print "Added CPoint at " & CStr(x) & " " & CStr(y - 0.15)
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

