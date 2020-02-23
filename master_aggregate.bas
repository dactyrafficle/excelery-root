Attribute VB_Name = "master_aggregate"
Sub aggro()

    Dim wb1 As Workbook
    Set wb1 = ActiveWorkbook
    
    Dim wb2 As Workbook
    Set wb2 = Workbooks.Add
    
    Dim y As Long
    y = 1
    
    Dim i As Long
    For i = 1 To wb1.Worksheets.Count
    
        Dim arr As Variant
        arr = Range(wb1.Worksheets(i).Range("a1"), wb1.Worksheets(i).Range("a1").End(xlToRight).End(xlDown)).Value
        
        wb2.Worksheets(1).Range("a" & y).Resize(UBound(arr, 1), UBound(arr, 2)).Value = arr
        
        y = y + UBound(arr, 1)
    
    Next i
    


End Sub
