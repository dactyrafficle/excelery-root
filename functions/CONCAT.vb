Option Explicit

'NEED FIXING
Function CONCAT(r As Range, ParamArray a() As Variant) As String

 'test for size
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

if y = ubound(arr,1) and x = ubound(arr,2) then
 str = str & arr(y,x)
else
   str = str & arr(y, x) & delimiter
end if
 
  Next x
 Next y

 CONCAT = str

End Function
