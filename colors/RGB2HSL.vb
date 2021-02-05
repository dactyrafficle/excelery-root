Public Function RGB2HSL(r_ As Double, g_ As Double, b_ As Double) As Variant

 'NORMALIZE
 Dim R As Double, G As Double, B As Double
 R = r_ / 255
 G = g_ / 255
 B = b_ / 255

 Dim arr(1 to 3) as variant
 
 'IF IS GREY; C=0
 If (R = G = B) Then
  arr(1) = 0 'h
  arr(2) = 0 's
  arr(3) = R 'l
  RGB2HSL = arr
  Exit Function
 End If

 Dim max As Double, min As Double
 max = WorksheetFunction.max(R, G, B)
 min = WorksheetFunction.min(R, G, B)

 Dim c As Double
 c = max - min

 Dim h As Double, h_ As Double

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
  
 'CONSIDER THESE LINES INSTEAD OF 32:45
 If (max = R) Then h_ = ((G - B) / c) - Int(((G - B) / c) / 6) * 6 'VBA MOD FN STINKS
 If (max = G) Then h_ = ((B - R) / c) + 2
 If (max = B) Then h_ = ((R - G) / c) + 4
    
 h = h_ * 60

 Dim l As Double
 l = (max + min) / 2

 Dim s As Double
 If (l <= 0.5) Then s = (max - min) / (max + min)
 If (l > 0.5) Then s = (max - min) / (2 - max - min)

 arr(1) = h
 arr(2) = s
 arr(3) = l

 RGB2HSL = arr

End Function
