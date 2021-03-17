Option Explicit

Public Function ASTACK(ParamArray arr() As Variant) As Variant
    
  'strip the arrays to content
    
  'think of mxn array
  'HOW BIG THE ARRAY NEEDS TO BE
  Dim m As Long, n As Long, count As Long
  count = 0
  For m = LBound(arr) To UBound(arr)
    Dim tarr As Variant
    tarr = arr(m)
    'WHY CANT I DO UBOUND(ARR(M),1) - LBOUND(ARR(M),1) ?
    count = count + UBound(tarr, 1) - LBound(tarr, 1) + 1
  Next m
  
  ReDim temp(1 To count)
  
  Dim y As Long
  y = 0
  For m = LBound(arr) To UBound(arr)
    tarr = arr(m)
    For n = LBound(tarr, 1) To UBound(tarr, 1)
      y = y + 1
      temp(y) = arr(m)(n, 1) 'arr IS A 1D-ARRAY, arr(m) IS A 2D-ARRAY
    Next n
  Next m
  
  ASTACK = temp

End Function


Sub asasa()

  Dim arr As Variant
  arr = ASTACK(Range("a1:a5"), Range("B1:B7"), Range("C1:C3"))

  Dim y As Long
  For y = 1 To UBound(arr, 1)
  
   Range("k" & y).Value = arr(y)
  Next y

End Sub
