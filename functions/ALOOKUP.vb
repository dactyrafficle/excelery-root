'the point of this function is that its like a vlookup but you can have it return a collection of results if there are multiple matches

Option Explicit

Function ALOOKUP(master As Variant, r As Range, x As Long, returnAsCollection As Boolean) As Variant

    'how can i make returnAsCollection optional?

    'save the range as a 2d array
    Dim arr As Variant
    arr = r.Value
    
    'loop over the first column, and look for master
    Dim i As Long
    Dim temp As New Collection 'to store successful results
    For i = LBound(arr, 1) To UBound(arr, 1)
        If arr(i, 1) = master Then
           temp.Add arr(i, x)
        End If
    Next i
    
    'for testing
    Dim j As Long
    For j = 1 To temp.Count
        Debug.Print temp(j)
    Next j
    
    If returnAsCollection Then
        ALOOKUP = temp
    Else
        ALOOKUP = temp(1)
    End If
    
End Function
