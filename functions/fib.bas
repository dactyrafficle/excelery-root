'Fibonacci using recursion
Function fib(n As Integer) As Long
  If n = 0 Then
    fib = 0
  ElseIf n = 1 Then
    fib = 1
  Else
    fib = fib(n - 1) + fib(n - 2)
  End If
End Function

'Fibonacci using arrays
Function fib_(n As Integer) As Long

  Dim fib As Long

  If n = 0 Then
    fib = 0
  End If

  If n = 1 Then
    fib = 1
  End If

  If n >= 2 Then
 
    Dim x() As Long
    ReDim x(0 To n)
    x(0) = 0
    x(1) = 1
 
    Dim i As Integer
    For i = 2 To n
      x(i) = x(i - 1) + x(i - 2)
    Next i
 
    fib = x(n)

  End If
 
  fib_ = fib

End Function
