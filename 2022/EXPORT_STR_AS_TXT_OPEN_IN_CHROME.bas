Public Sub EXPORT_STR_AS_TXT_OPEN_IN_CHROME(str As String)

  Dim fso As Scripting.FileSystemObject
  Set fso = New Scripting.FileSystemObject
  
  Dim FilePath As String, FileName As String
  FileName = "File-" & Timer * 1000 & ".txt"
  FilePath = "C:\Users\" & Environ("username") & "\Downloads\"
  
  Dim f As Scripting.TextStream
  Set f = fso.CreateTextFile(FilePath & FileName, 1, 0)

  f.WriteLine str
  
  'CLOSE FSO
  f.Close
  Set fso = Nothing
  
  Dim path As String, url As String
  path = "C:\Program Files (x86)\Google\Chrome\Application\Chrome.exe"
  url = "file:///C:/Users/" & Environ("username") & "/Downloads/" & FileName
  Shell (path & " -new-tab -url " & url)

End Sub