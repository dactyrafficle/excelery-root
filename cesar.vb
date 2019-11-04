Attribute VB_Name = "Module1"

Function cesar(str As String, shift As Integer) As String

    'arr = split(str, delimiter) [as an array] - no way to split into chars - so it stinks
    
    Dim bytes() As Byte
    bytes = StrConv(str, vbFromUnicode) 'an array of bytes
    
    'the range of acceptable chars will be unicode 32 to 126, for a range of 95 chars
    '26 lower, 26 upper, 10 digits, 33 spec chars = 95 chars
    Dim minCharNo, maxCharNo, charSpace As Integer
    minCharNo = 32
    maxCharNo = 126
    charSpace = maxCharNo - minCharNo + 1
    
    Dim output As String
    output = ""
    
    Dim i As Integer
    For i = 0 To UBound(bytes)
    
        Dim newCharNumber As Integer
        
        'vba has a buggy mod fn that return weird for negative values of a so i write my own, is easy enuf
    
        'c = mod(a, b) or c = (a)mod(b)
        Dim a, b, c As Integer
        a = Int((bytes(i) - minCharNo + shift)) 'convert range from 32-127 to 0-94, then shift
        b = charSpace
        c = a - Int(a / b) * b
        newCharNumber = c + minCharNo 'reconvert range to 32-127
        output = output & Chr(newCharNumber)
    
    Next i
    
    cesar = output

End Function

