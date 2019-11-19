Function xorShift(str1 As String, str2) As String

    'xorshift the 1st string by the values of the 2nd

    'convert str1 and str2 into byte arrays
    Dim msg() As Byte, pass() As Byte
    msg = StrConv(str1, vbFromUnicode)
    pass = StrConv(str2, vbFromUnicode)
    
    Dim output As String
    output = ""
    
    Dim i As Integer
    For i = LBound(msg) To UBound(msg)

        Dim x1, x2, b, y As Integer
        x1 = msg(i)

        b = UBound(pass) + 1
        x2 = pass(i Mod b)
    
    'so this does work, but i think 97 xor 97 is a bad result i think - how to fix
        Debug.Print x1 & " xor " & x2
        
        ' ie. 10 Xor 8 ->1010 Xor 1000 -> 0010 = 2
        y = x1 Xor x2
        
        output = output & Chr(y)
    
    Next i
    
    xorShift = output

End Function
