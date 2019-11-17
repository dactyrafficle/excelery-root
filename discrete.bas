Option Explicit

Public Function discrete(r1_ As Range, r2_ As Range) As String 'declare the output type; dont let excel pick for you

    'dont redefine r1_ and r2_; assign new temp variables
    Dim r1 As Range, r2 As Range
    Set r1 = r1_
    Set r2 = r2_

    'shrink column ranges
    Set r1 = Intersect(r1.Parent.UsedRange, r1)
    Set r2 = Intersect(r1.Parent.UsedRange, r2)

    'convert ranges to arrays
    Dim arr1 As Variant, arr2 As Variant
    arr1 = r1.Value
    arr2 = r2.Value

    Dim n As Long
    n = Application.Sum(arr2) 'sum total of 2nd array
    
    Dim x As Double
    x = n * Rnd() 'generate a random number 0 <= x < n

    Dim min As Double, max As Double
    min = 0
    max = 0

    Dim i As Long
    Dim output As String
    output = "error"
          
    'if arr2 has more numbers than strings in arr1, n will be big, and theres a chance we get an error
    For i = LBound(arr1, 1) To UBound(arr1, 1) 'where on the line x falls
    
        If max >= n Then
            Exit For
        End If
    
          If Not IsEmpty(arr1(i, 1)) And IsNumeric(arr2(i, 1)) Then 'make sure content in arr1, and number in arr2

            max = max + arr2(i, 1)
            If x >= min And x < max Then
                output = arr1(i, 1)
            End If
            min = min + arr2(i, 1)
        End If
    Next i
  
    discrete = output

End Function
