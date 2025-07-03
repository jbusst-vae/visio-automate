Attribute VB_Name = "lib"


Public  Sub test()
    Debug.Print "Start debug out"
    Debug.Print CStr(fileExists(ThisWorkbook.Path & "/vba-files/runVisio.bas"))
    Debug.Print CStr(fileExists(ThisWorkbook.Path & "/vba-files/runVisio"))
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