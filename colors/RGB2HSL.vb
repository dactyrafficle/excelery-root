Public Function RGB2HSL(r_ As Double, g_ As Double, b_ As Double) As Variant

 Dim R As Double, G As Double, B As Double
 R = r_ / 255
 G = g_ / 255
 B = b_ / 255

 If (R = G = B) Then
  'MEANS THE COLOR IS GREY
 End If

 Dim max As Double, min As Double
 max = WorksheetFunction.max(R, G, B)
 min = WorksheetFunction.min(R, G, B)

 Dim c As Double
 c = max - min

 Dim h As Double, h_ As Double

 'THIS MEANS THE COLOR IS GREY
 If (c = 0) Then
  h = 0
  Dim arr2(1 To 3) As Double
  arr2(1) = h
  arr2(2) = 0   'S = 0
  arr2(3) = max 'L = MAX = R = G = B
  rgb2hsl = arr2
  Exit Function
 End If

 'VBA MOD FN DOESNT HANDLE NEGATIVE VALUES WELL
 If (max = R) Then
  'h_ = ((g - b) / c) Mod 6 -> should work but vba mod fn stinks
  h_ = ((G - B) / c)
  h_ = h_ - Int(h_ / 6) * 6
 End If

 If (max = G) Then
  h_ = ((B - R) / c) + 2
 End If

 If (max = B) Then
  h_ = ((R - G) / c) + 4
 End If

 h = h_ * 60

 Dim l As Double
 l = (max + min) / 2

 Dim s As Double
 If (l <= 0.5) Then
  s = (max - min) / (max + min)
 Else
  s = (max - min) / (2 - max - min)
 End If

 Dim arr(1 To 6) As Variant
 arr(1) = h
 arr(2) = s
 arr(3) = l
 'arr(4) = max
 'arr(5) = min
 'arr(6) = c

 RGB2HSL = arr

End Function
