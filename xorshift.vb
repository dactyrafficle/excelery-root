Function xorShift(str1 As String, str2) As String

    'xorshift the 1st string by the values of the 2nd

    'convert str1 and str2 into byte arrays
    Dim msg() As Byte, pass() As Byte, out() As Byte
    msg = StrConv(str1, vbFromUnicode)
    pass = StrConv(str2, vbFromUnicode)
    out = StrConv(str1, vbFromUnicode)
    
    'Dim output As String
    'output = ""
    
    Dim i As Integer
    For i = LBound(msg) To UBound(msg)

        Dim x1, x2, b, y As Integer
        x1 = msg(i)

        b = UBound(pass) + 1
        x2 = pass(i Mod b)
        
        out(i) = x1 Xor x2
        
        'Debug.Print x1 & " xor " & x2; " -> " & out(i)
    
    Next i
    
    Debug.Print UBound(out) - LBound(out) + 1
    
    ' xorshift("abcde", "ab") produces a value with length 0. i want a blank with length 5
    xorShift = StrConv(out, vbUnicode)

End Function
