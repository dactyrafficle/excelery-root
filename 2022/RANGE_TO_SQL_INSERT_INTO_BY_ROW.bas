Attribute VB_Name = "xyz"
Option Explicit

Public Sub GET_SQL_INSERT_INTO(control As IRibbonControl)

    Call RANGE_TO_SQL_INSERT_INTO_BY_ROW

End Sub


Sub RANGE_TO_SQL_INSERT_INTO_BY_ROW()
 
  Dim doubleQuote As String
  doubleQuote = Chr(34)
 
  Dim r As Range
  
  If (IsEmpty(Range("a1"))) Then
   MsgBox "there is nthing here"
   Exit Sub
  End If
  
  Dim table_name As String
  table_name = Application.InputBox("TABLE NAME :")
  
  Set r = Range("a1", Range("a1").End(xlToRight).End(xlDown))
  
  Dim arr As Variant
  arr = r.Value
  
  Dim fso As Scripting.FileSystemObject
  Set fso = New Scripting.FileSystemObject
  
  Dim FilePath As String, FileName As String
  FileName = "File-" & Timer * 1000 & ".txt"
  FilePath = "C:\Users\" & Environ("username") & "\Downloads\"
  
  Dim f As Scripting.TextStream
  Set f = fso.CreateTextFile(FilePath & FileName, 1, 0)
  
  
  'GET THE ARRAY OF OBJECTS
   
  f.WriteLine ""
  
  Dim header As String
  header = header & "INSERT INTO " & table_name & " ("

  Dim x As Long
  For x = 1 To UBound(arr, 2)
    header = header & arr(1, x)
    
    If (x <> UBound(arr, 2)) Then
      header = header & ", "
    End If
    
    If (x = UBound(arr, 2)) Then
      header = header & ")"
    End If
    
  Next x
  header = header & ""
  
  
  Dim y As Long
  For y = 2 To UBound(arr, 1)

    
    Dim str As String
    str = header & " VALUES ("
    
    For x = 1 To UBound(arr, 2)
    

    
      str = str & doubleQuote & arr(y, x) & doubleQuote
    
      If (x <> UBound(arr, 2)) Then
        str = str & ","
      End If
    
    Next x
  
    str = str & ");"
    
    If (y <> UBound(arr, 1)) Then
      ' str = str & ","
    End If
    
    f.WriteLine str
  
  Next y
  
  f.WriteLine ""
  
  'CLOSE FSO
  f.Close
  Set fso = Nothing
  
  Dim path As String, url As String
  path = "C:\Program Files (x86)\Google\Chrome\Application\Chrome.exe"
  url = "file:///C:/Users/" & Environ("username") & "/Downloads/" & FileName
  Shell (path & " -new-tab -url " & url)

End Sub

