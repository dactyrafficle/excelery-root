Option Explicit

'accepts ranges and arrays
'isExact is only there to keep the same arg sequence as vlookup
'return as array is preferable over collection bc sum/count work on an array
Function ALOOKUP2(master As Variant, arr_ As Variant, x As Long, isExact As Boolean, returnAsArray As Boolean) As Variant
        
  'IF arr_ IS RANGE, SHRINK AND CONVERT TO ARR
  Dim arr As Variant
  If TypeName(arr_) = "Range" Then
    Dim r As Range
    Set r = Intersect(arr_.Parent.UsedRange, arr_)
    arr = r.Value
  Else
    arr = arr_
  End If

  'loop over the first column, and look for master
  Dim i As Long, count As Long, temp() As Variant 'to store results
  count = 0
  For i = LBound(arr, 1) To UBound(arr, 1)
    If arr(i, 1) = master Then
      count = count + 1
      ReDim Preserve temp(count)
      temp(count) = arr(i, x)
    End If
  Next i
    
  If returnAsArray Then
    ALOOKUP2 = temp
  Else
    ALOOKUP2 = temp(1)
  End If
  
End Function