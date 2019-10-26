Option Explicit

'this function accepts 2 ranges: r1 for the elements, r2 for the frequencies

Function rdiscrete(r1 As Range, r2 As Range)

  'testing
  'Sub abc()
  'Dim r1, r2 As Range
  'Set r1 = Range("a2:a4")
  'Set r2 = Range("b2:b4")

Dim n1, n2, n As Integer
n1 = r1.Count
n2 = r2.Count

If n1 > n2 Then
  n = n2
Else
  n = n1
End If

'there are n bins

Dim p As Integer
p = 0

Dim r As Range
For Each r In r2
    p = p + r.Value
Next r

'MsgBox p
'abc = p

Dim x() As Variant
ReDim x(1 To p)

Dim i, counter As Integer
counter = 0
For i = 1 To n
    Dim name As Variant
    name = r1(i)
    Dim freq As Integer
    freq = r2(i)
    Dim j As Integer
    For j = counter + 1 To counter + freq
        'this how many times we add name to the array
        x(j) = name
        Debug.Print (name)        
    Next j
    counter = counter + freq
Next i

'MsgBox "hi"

Dim y As Integer
y = Int(Rnd() * UBound(x) + 1)

Debug.Print (y)

rdiscrete = x(y)


End Function
