Attribute VB_Name = "Module1"

'imagine ws1 and ws2 contain data in cells a1:e1000 that you think should be the same, and you want to verify

Sub abc()

Dim t1 As Double
t1 = Timer

Dim sheet1, sheet2 As Worksheet
Set sheet1 = Sheets("Sheet1")
Set sheet2 = Sheets("Sheet2")

'save the data as 2d arrays so were not flipping back and forth from ws to ws
Dim arr1, arr2 As Variant
arr1 = sheet1.Range("a1:e1000").Value
arr2 = sheet2.Range("a1:e1000").Value

Dim x, y As Integer
For x = 1 To UBound(arr1, 2) 'for x
    For y = 1 To UBound(arr1, 1) 'for y
    
        If arr1(y, x) = arr2(y, x) Then
        
            'if the same, light green
            sheet1.Cells(y, x).Interior.Color = RGB(128, 255, 128)
        
        Else
        
            'if different, light red
            sheet1.Cells(y, x).Interior.Color = RGB(255, 102, 102)
        
        End If
    
    Next y
Next x

Debug.Print "time: " & Timer - t1

End Sub
