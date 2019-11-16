Option Explicit

Function discrete(r1 As Range, r2 As Range)

  'shrink column ranges to contiguous cells
  Set r1 = Range(r1(1), r1(1).End(xlDown))
  Set r2 = Range(r2(1), r2(1).End(xlDown))

  'convert ranges to arrays
  Dim arr1, arr2 As Variant
  arr1 = r1.Value
  arr2 = r2.Value

  Dim n As Long
  n = Application.Sum(arr2) 'sum total of 2nd array

  Dim x As Double
  x = n * Rnd() 'generate a random number 0 <= x < n

  Dim min, max As Double
  min = 0
  max = 0

  Dim i As Long
  For i = LBound(arr1, 1) To UBound(arr1, 1) 'where on the line x falls
      max = max + arr2(i, 1)
      If x >= min And x < max Then
          discrete = arr1(i, 1)
      End If
      min = min + arr2(i, 1)
  Next i

End Function
