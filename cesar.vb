'the start of a simple cesar shift; maybe focus on a range of first 0-127 chars from unicode
'no error for missing shift - needed
'how shall i do the reverse? a separate function or a parameter? 

Function cesar(str As String, shift As Integer) As String

    'arr = split(str, delimiter) [as an array] - no way to split into chars - so it stinks
    
    Dim bytes() As Byte
    bytes = StrConv(str, vbFromUnicode) 'an array of bytes
    
    Dim output As String
    output = ""
    
    Dim i As Integer
    For i = 0 To UBound(bytes)
    
        output = output & Chr(bytes(i) + shift)
    
    Next i
    
    cesar = output

End Function
