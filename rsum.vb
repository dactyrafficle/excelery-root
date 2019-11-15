'taking the intersection between the input range and the used range shrinks the range a lot
'it can make your code run 1000s of times faster

'also convert the range to an array
'itll also make it 1000x faster

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
