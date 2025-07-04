Attribute VB_Name = "lib"



Public Sub PrintArr(arr As Variant)
    Debug.Print "Arr: " & Join(arr, ", ")
End Sub

Public Sub Sleep(seconds As Double)
    Application.Wait Now + TimeSerial(0, 0, seconds)
End Sub

Function fileExists(ByVal path As String) As Boolean
    'Only works for a COMPLETE path
    'Returns TRUE if the provided path points to an existing file.
    'Returns FALSE if not existing, or if it's a folder
    fileExists = False
    Debug.Print path
    
    On Error Resume Next
    fileExists = ((GetAttr(path) And vbDirectory) <> vbDirectory)
End Function
