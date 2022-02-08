Attribute VB_Name = "abc"
Option Explicit

Public Sub GET_ARR_OF_OBJS_BY_ROW(control As IRibbonControl)

    Call RANGE_TO_ARR_OF_OBJS_BY_ROW

End Sub

Sub RANGE_TO_ARR_OF_OBJS_BY_ROW()
 
  Dim doubleQuote As String
  doubleQuote = Chr(34)
 
  Dim r As Range
  
  If (IsEmpty(Range("a1"))) Then
   MsgBox "there is nthing here"
   Exit Sub
  End If
  
  Set r = Range("a1", Range("a1").End(xlToRight).End(xlDown))
  'MsgBox r.Address
  
  Dim arr As Variant
  arr = r.Value
  
  Dim fso As Scripting.FileSystemObject
  Set fso = New Scripting.FileSystemObject
  
  Dim FilePath As String, FileName As String
  'Application.UserName v Environ("username")
  FileName = "File-" & Timer * 1000 & ".txt"
  FilePath = "C:\Users\" & Environ("username") & "\Downloads\"
  
  Dim f As Scripting.TextStream
  Set f = fso.CreateTextFile(FilePath & FileName, 1, 0)
  
  
  'GET THE ARRAY OF OBJECTS
   
  f.WriteLine "["
  
  Dim y As Long, x As Long
  For y = 2 To UBound(arr, 1)
  
    Dim str As String
    str = " {"
  
    For x = 1 To UBound(arr, 2)
    
      str = str & doubleQuote & arr(1, x) & doubleQuote & ":" & doubleQuote & arr(y, x) & doubleQuote
    
      If (x <> UBound(arr, 2)) Then
        str = str & ","
      End If
    
    Next x
  
    str = str & "}"
    
    If (y <> UBound(arr, 1)) Then
      str = str & ","
    End If
    
    f.WriteLine str
  
  Next y
  
  f.WriteLine "]"
  
  'CLOSE FSO
  f.Close
  Set fso = Nothing
  
  Dim path As String, url As String
  path = "C:\Program Files (x86)\Google\Chrome\Application\Chrome.exe"
  url = "file:///C:/Users/" & Environ("username") & "/Downloads/" & FileName
  Shell (path & " -new-tab -url " & url)

End Sub
