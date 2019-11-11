Sub copyDataToExcel()
    
    'create an instance of excel.exe
    Dim excel
    Set excel = CreateObject("excel.application")
    excel.Visible = True

    excel.workbooks.Add
    
    Dim data2 As Variant
    data2 = getTableDataAs2dArray(ActiveDocument.Tables(1))

    Dim z, n_items As Integer
    z = 0
    n_items = 0
    Dim x, y As Integer
    For y = 1 To UBound(data2, 1)
        'make sure the row has a non-zero qty
        If data2(y, 3) = 0 Or data2(y, 3) = "" Then
            'if qty is 0, increment z
            z = z + 1
        Else
            'if qty is non-zero, increment n_items & print the lines to excel via assignment
            n_items = n_items + 1
            For x = 1 To UBound(data2, 2)
                excel.Range("a1").Offset(y - 1 - z, x - 1).value = data2(y, x)
                'Debug.Print data2(y, x)
            Next x
        End If
    Next y
    
End Sub


'amazing, it works perfectly
Private Function getTableDataAs2dArray(tbl As Table) As Variant

    'Dim table1 As Table
    'Set table1 = ActiveDocument.Tables(1)

    Dim n_items, n_cols As Integer
    n_items = tbl.Rows.Count
    n_cols = tbl.Columns.Count

    Dim arr() As Variant
    ReDim arr(1 To n_items, 1 To n_cols)

    Dim x, y As Integer
    For y = 1 To tbl.Rows.Count
        For x = 1 To tbl.Columns.Count
            
            Dim r As Range
            Set r = tbl.Cell(y, x).Range
            r.End = r.End - 1
            
            Dim value As Variant
            value = r.Text
            
            arr(y, x) = value
        
        Next x
    Next y
    
    getTableDataAs2dArray = arr

End Function
