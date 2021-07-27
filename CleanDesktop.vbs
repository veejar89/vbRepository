Dim FSO
BasePath = ""'your base path here'
Set FSO = CreateObject("Scripting.FileSystemObject")
Set oFSO = FSO.GetFolder(BasePath)

For each file in oFSO.files
  If file.Attributes = ReadOnly Then
    file.Attributes = Normal
  End If
  fileExt = UCase(Split(file, ".")(1))
  DestPath = BasePath & fileExt & "/"
  If NOT FSO.FOlderExists(DestPath) Then
    FSO.CreateFolder DestPath
  End If
  If FSO.FileExists(file) Then
    If NOT FSO.FileExists(DestPath) Then
      FSO.MoveFile file, DestPath
    End If
  End If
Next
