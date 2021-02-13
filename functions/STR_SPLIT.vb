Option Explicit

Public Function STR_SPLIT(str As String, ParamArray delimiters() As Variant) As Variant

 'FIRST CHAR OF STR
 '0 : ignore
 '1 : +el to arr, +char to last el
 
 'ELSE
 '0 0 : ignore
 '1 0 : ignore
 '0 1 : +el to arr, +char to last el
 '1 1 : +char to last el

 Dim bytes() As Byte
 bytes = StrConv(str, vbFromUnicode) 'ARRAY OF BYTES
 
 Dim arr() As Variant, n As Long
 n = 0
 
 Dim y As Long, x As Long
 For y = LBound(bytes) To UBound(bytes)
 
 For x = LBound(delimiters) To UBound(delimiters)
  If (Chr(bytes(y)) = delimiters(x)) Then bytes(y) = 0
 Next x
  
  If (y = LBound(bytes)) Then 'IS FIRST CHAR
  
    If (bytes(y) <> 0) Then
      n = n + 1
      ReDim Preserve arr(1 To n)
      arr(n) = arr(n) & Chr(bytes(y))
    End If
  
  Else 'NOT FIRST CHAR
  
    '0 1
    If (bytes(y) <> 0 And bytes(y - 1) = 0) Then
      n = n + 1
      ReDim Preserve arr(1 To n)
      arr(n) = arr(n) & Chr(bytes(y))
    End If
  
    '1 1
    If (bytes(y) <> 0 And bytes(y - 1) <> 0) Then arr(n) = arr(n) & Chr(bytes(y))
  
  End If
 Next y

 STR_SPLIT = arr

End Function
