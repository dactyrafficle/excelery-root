Option Explicit

Function rcharstring(n As Long) As String

    Dim str As String
    str = ""
    
    Dim min As Long, max As Long
    min = 97
    max = 125
    
    Dim i As Long, Dim x as Long
    For i = 1 To n
        x = Int(min + Rnd() * (max - min + 1))
        str = str & Chr(x)
    Next i
    
    rcharstring = str

End Function
