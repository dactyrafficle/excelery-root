'in vba, if you use a 2d array, you can only resize via redim preserve the last dimension
'that makes it hard for me to store data if i know how many columns i want but not how many rows
'so in this case i would rather do like i would in javascript, and make an array of arrays
'this way, the main array is a 1d array, and its elements are all 1d arrays
'so its like it was a 2d array, but really, its an array of arrays

Sub abc()

Dim r As Range
Set r = Range("a1:c11") 'a: number, b: price, c: qty

Dim arr() As Variant
Dim count As Long
count = 0

For Each Row In r.Rows

    If IsNumeric(Row.Cells(1, 3).Value) And Row.Cells(1, 3).Value > 0 Then
   
        count = count + 1
        
        ReDim Preserve arr(1 To count)
        
        Dim arr2(1 To 3) As Variant
        arr2(1) = Row.Cells(1).Value
        arr2(2) = Row.Cells(2).Value
        arr2(3) = Row.Cells(3).Value
        
        arr(count) = arr2
    
    End If

Next Row

Dim i As Integer
For i = LBound(arr) To UBound(arr)
    
    Debug.Print arr(i)(1) & ":" & arr(i)(2) & "x" & arr(i)(3)

Next i


End Sub
