Function xorShift(str As String) As String

    'ideally we would use a stream of binarys to shift this but here we'll just setting for random

    Dim bytes() As Byte
    bytes = StrConv(str, vbFromUnicode) 'an array of bytes
    
    Dim output As String
    output = ""
    
    Dim i As Integer
    For i = 0 To UBound(bytes)
    
        Dim x1, x2, y As Integer
        x1 = bytes(i)
        x2 = Int(Rnd() * 128)
        
        ' ie. 10 Xor 8 -> 1010 Xor 1000 -> 0010 = 2
        y = x1 Xor x2
        
        output = output & Chr(y)
    
    Next i
    
    xorShift = output

End Function
