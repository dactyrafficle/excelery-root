Option Explicit

'accepts ranges and arrays
'isExact is only there to keep the same arg sequence as vlookup
Function ALOOKUP(master As Variant, a As Variant, x As Long, isExact As Boolean, returnAsCollection As Boolean) As Variant

    'if a is a range, shrink before converting to an array
    Dim arr As Variant
    If TypeName(a) = "Range" Then
        Dim r As Range
        Set r = Intersect(a.Parent.UsedRange, a)
        arr = r.Value
        Debug.Print UBound(arr, 1)
    Else
        arr = a
    End If

    
    'loop over the first column, and look for master
    Dim i As Long
    Dim temp As New Collection 'to store successful results
    For i = LBound(arr, 1) To UBound(arr, 1)
        If arr(i, 1) = master Then
           temp.Add arr(i, x)
        End If
    Next i
    
    If returnAsCollection Then
        ALOOKUP = temp
    Else
        ALOOKUP = temp(1)
    End If
    
End Function
