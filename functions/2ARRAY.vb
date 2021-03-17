Function abc(x As Variant) As Variant

  '[yes] single cell: ie. r = Range("a1")
  '[yes] many cells: ie. r = Range("a1:a100")
  '[yes] 1d-array: ie. arr = Array(1,2,3)
  '[yes] 2d-array: ie. arr = Array(Array(1,2,3),Array(4,5,6),Array(7,8,9))
  '[yes] single value: ie. x=5
  
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

Sub Test()

  Dim x As Long, y As String
  x = 5
  y = "this"
  Debug.Print abc(x) 'returns 1
  Debug.Print abc(y) 'returns 1
  
  Dim arr1 As Variant, arr2 As Variant
  arr1 = Array(1, 2, 15)
  arr2 = Array(Array(1, 2, 3, 7), Array(1, 2))

  Debug.Print abc(arr1) 'returns 3
  Debug.Print abc(arr2) 'returns 2

End Sub
