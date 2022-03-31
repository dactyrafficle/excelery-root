Attribute VB_Name = "sub_general_Rng2ArrOfObjs"
Option Explicit

Public Sub general_Rng2ArrOfObjs(control As IRibbonControl)

    Call Rng2ArrOfObjs

End Sub

Sub Rng2ArrOfObjs()
 
  Dim doubleQuote As String, single_space As String
  doubleQuote = Chr(34)
  single_space = " "
 
  Dim r As Range
  
  If (IsEmpty(Range("a1"))) Then
   MsgBox "there is nthing here"
   Exit Sub
  End If
  
  Set r = Range("a1", Range("a1").End(xlToRight).End(xlDown))

  'COMMENT THIS OUT
  'r.Interior.Color = RGB(255, 230, 255)
  'MsgBox r.Address
  
  'CONVER TO ARR
  Dim arr As Variant
  arr = r.Value
  
  Dim fso As Scripting.FileSystemObject
  Set fso = New Scripting.FileSystemObject
  
  'SET WHERE THE FILE WILL BE SAVED
  Dim FilePath As String, FileName As String
  FileName = "File-" & Timer * 1000 & ".txt"
  FilePath = "C:\Users\" & Environ("username") & "\Downloads\"
  
  Dim f As Scripting.TextStream
  Set f = fso.CreateTextFile(FilePath & FileName, 1, 0)
  
  'INITIALIZE THE ARR STRING
  Dim JSON_ARR_STR As String
  JSON_ARR_STR = "["
  JSON_ARR_STR = JSON_ARR_STR & vbNewLine
  

  
  'SO FAR, SO GOOD
  
  Dim y As Long, x As Long
  For y = 2 To UBound(arr, 1)
  
    JSON_ARR_STR = JSON_ARR_STR & single_space & "{" & vbNewLine
    
    For x = 1 To UBound(arr, 2)
    
      JSON_ARR_STR = JSON_ARR_STR & single_space & single_space & doubleQuote & arr(1, x) & doubleQuote & ":" & doubleQuote & arr(y, x) & doubleQuote
    
      If (x <> UBound(arr, 2)) Then
        JSON_ARR_STR = JSON_ARR_STR & ","
      End If
    
    JSON_ARR_STR = JSON_ARR_STR & vbNewLine
    
    Next x
    
    JSON_ARR_STR = JSON_ARR_STR & single_space & "}"
    
    If (y <> UBound(arr, 1)) Then
      JSON_ARR_STR = JSON_ARR_STR & "," & vbNewLine
    End If
  
  Next y
  
  'CLOSE THE JS ARRAY
  JSON_ARR_STR = JSON_ARR_STR & vbNewLine & "]"
  
  'PRINT
  Debug.Print JSON_ARR_STR
  f.WriteLine JSON_ARR_STR
  
  'CLOSE FSO
  f.Close
  Set fso = Nothing
  
  Dim path As String, url As String
  path = "C:\Program Files (x86)\Google\Chrome\Application\Chrome.exe"
  url = "file:///C:/Users/" & Environ("username") & "/Downloads/" & FileName
  Shell (path & " -new-tab -url " & url)

End Sub
