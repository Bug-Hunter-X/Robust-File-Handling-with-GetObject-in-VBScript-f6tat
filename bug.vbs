Function GetObject(path)
  On Error Resume Next
  Set GetObject = GetObject(path)
  If Err.Number <> 0 Then
    Err.Clear
    Set GetObject = Nothing
  End If
End Function

' Example usage
Dim obj
Set obj = GetObject("C:\\some\\file.txt")
If obj Is Nothing Then
  MsgBox "File not found!"
Else
  ' Process the object
End If