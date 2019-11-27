Attribute VB_Name = "Module1"

'this function is for returning a random string of characters

Function rcharstring(n As Integer) As String

    Dim str As String
    str = ""
    
    Dim min, max As Integer
    min = 48
    max = 125
    
    Dim i As Integer
    
    For i = 1 To n
    
        Dim x As Integer
        x = Int(min + Rnd() * (max - min + 1))
        str = str & Chr(x)
    
    Next i
    
    rcharstring = str

End Function
