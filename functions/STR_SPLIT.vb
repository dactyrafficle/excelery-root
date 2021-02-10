Option Explicit

 

Public Function STR_SPLIT(str As String) As Variant

 

Dim bytes() As Byte

bytes = StrConv(str, vbFromUnicode) 'ARRAY OF BYTES

 

'char(32) is blank space

'char(10) is newline char

 Dim arr() As Variant

Dim n As Long

n = 0

 Dim i As Long

For i = LBound(bytes) To UBound(bytes)

  'IS FIRST CHARACTER?

  If (i = LBound(bytes)) Then

 

    If (bytes(i) <> 32 And bytes(i) <> 10) Then

     'add element to array

     n = n + 1

     ReDim Preserve arr(1 To n)

    arr(n) = arr(n) & Chr(bytes(i))

    End If

  

  Else

  'NOT FIRST ELEMENT

  

   

   'THE PREVIOUS ELEMENT IS NOT SOLID

   If (bytes(i - 1) = 32 Or bytes(i - 1) = 10) Then

   

    'CURRENT IS SOLID

     If (bytes(i) <> 32 And bytes(i) <> 10) Then

      'ADD ELEMENT TO ARRAY

      n = n + 1

      ReDim Preserve arr(1 To n)

      arr(n) = arr(n) & Chr(bytes(i))

     Else

      'CURRENT IS NOT SOLID

      'SKIP

     End If


   Else

   'PREVIOUS ELEMENT IS SOLID

    'CURRENT IS SOLID

     If (bytes(i) <> 32 And bytes(i) <> 10) Then

      'ADD CHARACTER TO LAST ELEMENT OF ARRAY

      arr(n) = arr(n) & Chr(bytes(i))

     Else

     'CURRENT IS NOT SOLID

      'SKIP

     End If


   End If

  End If

 Next i

STR_SPLIT = arr

End Function
