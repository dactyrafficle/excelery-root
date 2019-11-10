Option Explicit

Function rsum(r1 As Range) As Double

    Set r1 = Intersect(r1.Parent.UsedRange, r1)
    
    Dim a, arr As Variant
    arr = r1.Value
    
    If Not r1 Is Nothing Then
        For Each a In arr
         rsum = rsum + a
        Next a
    End If

End Function
