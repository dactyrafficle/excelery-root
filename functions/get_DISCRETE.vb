Option Explicit

Function get_DISCRETE(r As Range)

 Dim arr() As Variant
 arr = r.VALUE

 Dim count As Long
 count = 0
 
 Dim y As Long
 For y = LBound(arr, 1) To UBound(arr, 1)
  count = count + arr(y, 2)
 Next y

 'NOW THAT WE HAVE THE TOTAL COUNT, WE CAN ASSIGN PROPORTIONS
 Debug.Print (count)


 Dim output As Variant
 output = arr(1, 1)
 
 Dim score As Double
 score = Rnd()
 
 Dim cdf As Double
 cdf = 0
 
 For y = LBound(arr, 1) To UBound(arr, 1)
 
   cdf = cdf + (arr(y, 2) / count)   'ADD THE CURRENT PROPORTION TO THE CDF
   
   If (score < cdf) Then
     output = arr(y, 1)
     Exit For
   End If

 Next y
 
 get_DISCRETE = output

End Function


