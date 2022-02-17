Private Function RANGE_TO_SQL_INSERT_INTO_STRING(table_name As String)

  Dim doubleQuote As String, newLineChar As String
  doubleQuote = Chr(34)
  newLineChar = Chr(10) '& Chr(13)

  Dim r As Range
  Set r = Range("a1", Range("a1").End(xlToRight).End(xlDown))
  
  Dim arr As Variant
  arr = r.Value

  'HEADER
  Dim header As String
  header = header & "INSERT INTO " & table_name & " ("
  
  Dim output As String
  
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
  
  'CONTENTS
  Dim y As Long
  For y = 2 To UBound(arr, 1)

    Dim record As String
    record = header & " VALUES ("
    
    For x = 1 To UBound(arr, 2)
    
      record = record & doubleQuote & arr(y, x) & doubleQuote
    
      If (x <> UBound(arr, 2)) Then
        record = record & ","
      End If
    
    Next x
  
    record = record & ");"
    
    'NEWLINE CHAR
    output = output & record & newLineChar
  
  Next y
  
  RANGE_TO_SQL_INSERT_INTO_STRING = output
  
End Function