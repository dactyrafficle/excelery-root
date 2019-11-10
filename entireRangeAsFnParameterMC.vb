Sub abc()
    Dim i As Integer
    For i = 1 To 2000
        If Rnd > 0.75 Then
            Cells(i, 1).Value = WorksheetFunction.RandBetween(100, 999)
        Else
            Cells(i, 1).Value = ""
        End If
    Next i
End Sub

Sub xFNA()

    Dim t As Double
    t = Timer
    
    Dim x As Double
    x = 0
    
    Dim i As Long
    For i = 1 To 100
        
        x = x + fnb(Range("a:a"))
        
    Next i

    Debug.Print "time: " & Timer - t

End Sub

Function fna(r1 As Range) As Double

    Dim arr As Variant
    arr = r1.Value

    For Each A In arr
        fna = fna + A
    Next A
    
End Function

Function fnb(r1 As Range) As Double

    Set r1 = Intersect(r1.Parent.UsedRange, r1)
    
    Dim arr As Variant
    arr = r1.Value

    If Not r1 Is Nothing Then
        For Each A In arr
         fnb = fnb + A
        Next A
    End If

End Function
