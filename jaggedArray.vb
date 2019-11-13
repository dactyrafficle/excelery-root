Sub abc()

'important to use lbound and ubound and remember ubound(arr,1) is y, and ubound(arr,2) is x, which is weird

Dim arr As Variant
ReDim arr(1 To 3, 1 To 5)
    
Dim y, x As Long
For y = 1 To UBound(arr, 1)

    For x = 1 To UBound(arr, 2)
    
        If x = 2 Then
        
            Dim arr2 As Variant
            ReDim arr2(1 To 4)
            Dim z As Long
            For z = 1 To UBound(arr2)
            
                arr2(z) = Int(Rnd() * 100)
            
            Next z
            
            arr(y, x) = arr2
            
        
        Else

            arr(y, x) = Int(Rnd() * 100)
        
        End If
    
    Next x

Next y


'print

Dim shift, counter As Long
shift = 0
For y = 1 To UBound(arr, 1)

    counter = 0

    For x = 1 To UBound(arr, 2)
    
        If x = 2 Then
        
            For z = 1 To UBound(arr2)
            
                Range("a1").Offset(y - 1 + shift + counter, x - 1).Value = arr2(z)
                
                counter = counter + 1 'even if its the last one
            
            Next z
            
        Else
    
            Range("a1").Offset(y - 1 + shift, x - 1).Value = arr(y, x)
        
        End If
        
    Next x
    
    shift = shift + counter

Next y


End Sub
