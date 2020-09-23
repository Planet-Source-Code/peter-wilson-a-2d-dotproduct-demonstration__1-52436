Attribute VB_Name = "Module1"
Option Explicit

' Shellexecute API to launch a file, or fire up a web browser etc
Public Const SW_SHOWNORMAL = 1
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

' This is my custom data type.
Public Type mdrVector
    x As Single
    Y As Single
End Type

Public Sub LaunchAnyFile(p_objForm As Form, p_strFileName As String)

    On Error GoTo errTrap
    
    Dim lngReturnValue As Long
    
    lngReturnValue = ShellExecute(p_objForm.hwnd, "open", p_strFileName, vbNull, vbNull, SW_SHOWNORMAL)
        
    Exit Sub
errTrap:
    MsgBox Err.Number & " - " & Err.Description, vbExclamation
    
End Sub


Public Function DotProduct(VectorU As mdrVector, VectorV As mdrVector) As Single

    ' Determines the dot-product of two vectors.
    DotProduct = (VectorU.x * VectorV.x) + (VectorU.Y * VectorV.Y)
    
End Function


Public Function Vec3Length(V1 As mdrVector) As Single

    ' Returns the length of a vector.
    Vec3Length = Sqr((V1.x ^ 2) + (V1.Y ^ 2))
    
End Function

Public Function Vec3Normalize(V1 As mdrVector) As mdrVector

    ' Returns the normalized version of a vector.
    
    Dim sngLength As Single
    
    sngLength = Vec3Length(V1)
    
    If sngLength = 0 Then sngLength = 1
    
    Vec3Normalize.x = V1.x / sngLength
    Vec3Normalize.Y = V1.Y / sngLength
    
End Function
