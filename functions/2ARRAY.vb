Function abc(x As Variant) As Variant

  'single cell: ie. r = Range("a1")
  'many cells: ie. r = Range("a1:a100")
  '1d-array: ie. arr = Array(1,2,3)
  '2d-array: ie. arr = Array(Array(1,2,3),Array(4,5,6),Array(7,8,9))
  'single value: ie. x=5
  
 'Debug.Print TypeName(x)
 'Debug.Print VarType(x)
 
 Dim arr As Variant
 
 If IsArray(x) Then
  arr = x
 Else
  arr = Array(x)
 End If

 Dim y As Long
 y = UBound(arr, 1) - LBound(arr, 1) + 1

 abc = y
 
End Function
