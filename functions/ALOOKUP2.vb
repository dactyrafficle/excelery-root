Option Explicit

'arr_ can be a range or arrays
'isApprox is only there to keep the same arg sequence as vlookup
'returnAsArray is preferable over collection bc SUM(), COUNT() and INDEX() work on arrays
Public Function ALOOKUP2(lookup_value As Variant, arr_ As Variant, x As Long, Optional isApprox As Boolean = False, Optional returnFirstMatch As Boolean = False) As Variant
        
  'IF arr_ IS RANGE, SHRINK AND CONVERT TO ARR
  Dim arr As Variant
  If TypeName(arr_) = "Range" Then
    Dim r As Range
    Set r = Intersect(arr_.Parent.UsedRange, arr_)
    arr = r.Value
  Else
    arr = arr_
  End If

  'LOOP OVER THE FIRST COLUMN, AND LOOK FOR lookup_value
  Dim i As Long, n As Long 'TO STORE THE NUMBER OF HITS
  Dim temp() As Variant 'AN ARRAY TO STORE RESULTS
  n = 0
  For i = LBound(arr, 1) To UBound(arr, 1) 
    If arr(i, 1) = lookup_value Then
      ReDim Preserve temp(n) 'RESIZE THE ARRAY TO n
      temp(n) = arr(i, x)
      n = n + 1 'INCREMENT n
    End If
  Next i
    
  If returnFirstMatch Then
    ALOOKUP2 = temp(0) 'RETURN THE FIRST HIT
  Else
    ALOOKUP2 = temp 
  End If
  
End Function
