Option Explicit

'a = 0
'b = 1
'z = 25
'aa = 26
'zz = 701
'abc = 730

Public Function get_global_index(ByVal str As String)

    'ALL LOWER CASE
    str = LCase(str)

    Dim bytes() As Byte
    bytes = StrConv(str, vbFromUnicode)

    'WITH VARIANT ARRAY ONLY
    Dim letters() As Variant
    letters() = Array("a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z")
    
    
    Dim output_a As String
    Dim output_b As Long

    Dim n As Long
    n = UBound(bytes) - LBound(bytes) + 1

    Dim x As Long
    For x = 0 To n - 1
    
        output_a = output_a & " + " & (bytes(x) - 97 + 1) & "x" & WorksheetFunction.Power(26, n - x - 1)
        output_b = output_b + (bytes(x) - 97 + 1) * WorksheetFunction.Power(26, n - x - 1)
        
    Next x
    
    'get_global_index = output_a
    get_global_index = output_b - 1

End Function
