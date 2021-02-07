Option Explicit

Public Function ALOOKUP3(ByVal lookup_arr_in As Variant, ByVal search_arr_in As Variant, return_col_index As Long, Optional isApprox As Boolean = False, Optional returnFirstMatch As Boolean = False) As Variant

  Debug.Print "TypeName(lookup_arr_in) = " & TypeName(lookup_arr_in)
  Dim lookup_arr() As Variant
  If TypeName(lookup_arr_in) = "Range" Then lookup_arr = Intersect(lookup_arr_in.Parent.UsedRange, lookup_arr_in).Value 'results in 2d-array
  If TypeName(lookup_arr_in) = "Variant()" Then lookup_arr = lookup_arr_in 'may result in 1d-array

  Dim z As Long
  'For z = LBound(lookup_arr) To UBound(lookup_arr)
  ' Debug.Print lookup_arr(z)
  'Next z
  
  Dim search_arr As Variant
  If TypeName(search_arr_in) = "Range" Then search_arr = Intersect(search_arr_in.Parent.UsedRange, search_arr_in).Value 'results in 2d-array
  If TypeName(search_arr_in) = "Variant()" Then search_arr = search_arr_in


  'LOOP OVER THE FIRST COLUMN, AND LOOK FOR lookup_value
  
  Dim y As Long, x As Long, n As Long
  Dim temp() As Variant     'TO STORE RESULTS
  n = 1
  For y = LBound(search_arr, 1) To UBound(search_arr, 1)
  
   For x = LBound(lookup_arr, 1) To UBound(lookup_arr, 1)
  
    If search_arr(y, 1) = lookup_arr(x, 1) Then
      ReDim Preserve temp(1 To n) 'RESIZE
      temp(n) = search_arr(y, return_col_index)
      n = n + 1     'INCREMENT n
    End If
  
   Next x
  Next y
    
  If returnFirstMatch Then
    ALOOKUP3 = temp(1) 'RETURN THE FIRST HIT
  Else
    ALOOKUP3 = temp
  End If
  
End Function


