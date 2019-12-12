Attribute VB_Name = "NewMacros"

Sub copyData2Excel()


    'create an instance of excel.exe
    Dim excel
    Set excel = CreateObject("excel.application")
    excel.Visible = True

    excel.workbooks.Add

    'how many sections?
    
    Dim n_rows As Long
    n_rows = 0

Dim m As Long
m = ActiveDocument.Sections.Count
Debug.Print m

'loop over the sections
Dim i As Long
For i = 1 To m

    'parse the invoice number
    Dim invoice_ As Variant, invoice As Variant
    invoice_ = ActiveDocument.Sections(i).Headers(1).Range.Tables(1).Cell(1, 2)

    Dim a As Long
    a = InStr(invoice_, "INVOICE:") + Len("INVOICE:") + 1

    invoice = "'" & Mid(invoice_, a, InStr(invoice_, "Shipment") - a)
    
    ' i need to PARSE THIS
    Dim order_ As Variant, order As Variant
    order_ = ActiveDocument.Sections(i).Headers(1).Range.Tables(2).Cell(7, 1)
    order = "'" & Mid(order_, 1, Len(order_) - 1)
    
    'MsgBox order

    'ActiveDocument.Sections(i).Range.Select
    
    'Dim order As Variant
    'order = ActiveDocument.Sections(i).Headers(1).Range.Tables(2).Cell(7, 1)
    
    'how many tables in this range?
    
    Dim n As Long
    n = ActiveDocument.Sections(i).Range.Tables.Count
    Debug.Print n
    
            'let us take the first table from this range
        
            Dim data2 As Variant
            data2 = getTableDataAs2dArray(ActiveDocument.Sections(i).Range.Tables(1))
            
            Dim x, y As Long
            For y = 1 To UBound(data2, 1)
                excel.Range("a1").Offset(n_rows, 0).value = order
                excel.Range("a1").Offset(n_rows, 1).value = invoice
                For x = 1 To UBound(data2, 2)
                    excel.Range("a1").Offset(n_rows, x + 1).value = data2(y, x)
                Next x
                n_rows = n_rows + 1
            Next y


Next i



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


