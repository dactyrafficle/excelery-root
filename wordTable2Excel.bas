'this is a custom function that i am using to extract data from invoices in .docx format
Sub copyData2Excel()

	'create an instance of excel.exe
	Dim excel
	Set excel = CreateObject("excel.application")
	excel.Visible = True

	'create a wb
	excel.workbooks.Add

	'in my case, im interested in invoices
	'one word doc can contain many invoices
	'each invoice is a section
	Dim m As Long
	m = ActiveDocument.Sections.Count
	
	'this variable will store the number of lines in the excel doc were copying the data to
	Dim n_rows As Long
	n_rows = 0

	'loop over the individual sections (invoices)
	Dim i As Long
	For i = 1 To m

		'parse the invoice number
		Dim invoice_ As Variant, invoice As Variant
		invoice_ = ActiveDocument.Sections(i).Headers(1).Range.Tables(1).Cell(1, 2)
		Dim a As Long
		a = InStr(invoice_, "INVOICE:") + Len("INVOICE:") + 1
		invoice = Mid(invoice_, a, InStr(invoice_, "Shipment") - a)

		'parse the order number
		Dim order_ As Variant, order As Variant
		order_ = ActiveDocument.Sections(i).Headers(1).Range.Tables(2).Cell(7, 1)
		order = Mid(order_, 1, Len(order_) - 1)
    
		'parse the customer number
		Dim customer_ As Variant, customer As Variant
		customer_ = ActiveDocument.Sections(i).Headers(1).Range.Tables(2).Cell(7, 4)
		customer = Mid(customer_, 1, Len(customer_) - 1)

		'the body of each section contains 2 tables: the first is the data, the 2nd is a summary
		Dim n As Long
		n = ActiveDocument.Sections(i).Range.Tables.Count
    
		'lets go after the first
		Dim data2 As Variant
		data2 = getTableDataAs2dArray(ActiveDocument.Sections(i).Range.Tables(1))
					
		Dim x as Long, y As Long
		For y = 1 To UBound(data2, 1)
		
			excel.Range("a1").Offset(n_rows, 0).value = "'" & customer
			excel.Range("a1").Offset(n_rows, 1).value = "'" & order
			excel.Range("a1").Offset(n_rows, 2).value = "'" & invoice
			
			For x = 1 To UBound(data2, 2)
			
				excel.Range("a1").Offset(n_rows, x + 2).value = data2(y, x)
				
			Next x
			n_rows = n_rows + 1
			
		Next y
			
	Next i

End Sub


'this function will return a table in word as a 2d array
Private Function getTableDataAs2dArray(tbl As Table) As Variant

    Dim n_rows as Long, n_cols As Long
    n_rows = tbl.Rows.Count
    n_cols = tbl.Columns.Count

    Dim arr() As Variant
    ReDim arr(1 To n_rows, 1 To n_cols)

    Dim x, y As Long
    For y = 1 To n_rows
        For x = 1 To n_cols
            
            Dim r As Range
            Set r = tbl.Cell(y, x).Range
            r.End = r.End - 1 'this is bc the doc im working with has a spleen
            
            Dim value As Variant
            value = r.Text
            
						'store the value of the cell in the array
            arr(y, x) = value
        
        Next x
    Next y
    
    getTableDataAs2dArray = arr

End Function