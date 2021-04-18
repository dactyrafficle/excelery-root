Option Explicit

Function DISCRETE(ParamArray x() As Variant)

 Dim n As Long, i As Long
 n = UBound(x) - LBound(x) + 1
 i = Int(Rnd() * n)
 
 DISCRETE = x(i)

End Function
