
Function getCellColorInfo(rng As Range) As Variant

 'VBA STORES COLOR AS A 3-BYTE NUMBER IN BGR SEQUENCE
 'SO REALLY, ITS LIKE A NUMBER IN BASE 256
 'IN THE SAME WAY THAT 101 = 5 IN BASE 2
 'HERE, MAKE SURE YOU REMOVE THE INFLUENCE OF THE EARLIER SEQUENCE BEFORE CALCULATING G OR B
 'FOR EXAMPLE, IF R=254, DOING C/256 WILL YIELD A NUMBER WITH A DECIMAL REPRESENTING 255/256
 
 Dim c As Double
 c = rng.Interior.Color
 
 Dim R As Double, G As Double, B As Double
 R = c Mod 256
 G = ((c - R) / 256) Mod 256
 B = ((c - G * 256 - R) / 65536) Mod 256
 
 Dim arr(1 To 3) As Variant
 arr(1) = R
 arr(2) = G
 arr(3) = B

 getCellColorInfo = arr

End Function
