Attribute VB_Name = "MyRibbon"

'namespace=vba-files/ribbons



Public Sub btn1(ByRef control As Office.IRibbonControl)
    runVisio.OpenVisio
End Sub



Public Sub btn2(ByRef control As Office.IRibbonControl)
    MsgBox "Button 2",,"title" 
End Sub



Public Sub btn3(ByRef control As Office.IRibbonControl)
    MsgBox "Button 3",,"title" 
End Sub