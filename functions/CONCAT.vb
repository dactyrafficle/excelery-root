Option Explicit

'NEED FIXING
Function CONCAT(r As Range, ParamArray a() As Variant) As String

 Dim arr As Variant
 arr = r.Value
 
 Dim str As String
 str = ""
 
 Dim delimiter As String
 delimiter = ""
 Dim z As Long
 For z = LBound(a) To UBound(a)
   delimiter = delimiter & a(z)
 Next z
 
 Dim y As Long, x As Long
 For y = LBound(arr, 1) To UBound(arr, 1)
  For x = LBound(arr, 2) To UBound(arr, 2)
 
   str = str & arr(y, x) & delimiter
 
  Next x
 Next y

 CONCAT = str

End Function
