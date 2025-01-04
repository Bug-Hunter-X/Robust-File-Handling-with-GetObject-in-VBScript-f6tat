Function GetObjectSafe(path)
  Dim objFSO, fileExists
  Set objFSO = CreateObject("Scripting.FileSystemObject")
  fileExists = objFSO.FileExists(path)
  Set objFSO = Nothing
  If fileExists Then
    On Error Resume Next
    Set GetObjectSafe = GetObject(path)
    If Err.Number <> 0 Then
      Err.Clear
      Set GetObjectSafe = Nothing
      MsgBox "Error accessing file: " & Err.Description, vbCritical
    End If
    On Error GoTo 0
  Else
    MsgBox "File not found: " & path, vbExclamation
    Set GetObjectSafe = Nothing
  End If
End Function

' Example Usage
Dim obj
Set obj = GetObjectSafe("C:\\some\\file.txt")
If obj Is Nothing Then
  ' Handle the case where the file doesn't exist or there was an error
Else
  ' Process the object here
End If