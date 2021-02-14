Public Function FILTER_ARR(input_arr As Variant, col_index As Long, filter_value As Variant) As Variant

 'INPUT ARRAY ROWS
 Dim n1 As Long, n2 As Long
 n1 = LBound(input_arr, 2)
 n2 = UBound(input_arr, 2)
 
 'OUTPUT ARRAY ROWS
 Dim m1 As Long, m2 As Long, count As Long
 m1 = LBound(input_arr, 1)
 m2 = LBound(input_arr, 1)
 count = 0

 Dim temp_arr() As Variant
 Dim y As Long, x As Long

 For y = m1 To UBound(input_arr, 1)
  If (input_arr(y, col_index) = filter_value) Then

   count = count + 1
   m2 = m2 + count
   ReDim Preserve temp_arr(n1 To n2, m1 To m2)
   For x = n1 To n2
     temp_arr(x, count) = input_arr(y, x)
    Next x
  End If
 Next y

 FILTER_ARR = Application.Transpose(temp_arr)

End Function
