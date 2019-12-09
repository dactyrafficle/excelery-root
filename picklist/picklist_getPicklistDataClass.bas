Attribute VB_Name = "picklist_getPicklistDataClass"

Sub getPicklistData()

    Dim str As String
    
    str = Application.InputBox("type 'ok, get data' to run")

    If str = "ok, get data" Then
    
        Call removeFormattingActions
        
        Dim order As New salesOrder
        Set order = getPicklistDataActions
        
        'test the method
        order.printHeader
    
    Else
    
        'do nothing
        
    End If

End Sub




Private Function getPicklistDataActions()

    ' order -> header, body
    ' header -> billTo, shipTo, summary
    ' body > data -> dataLines [data is a collection of dataLine objects]
    
    Dim order As New salesOrder
    
    'get the header data

    order.header.summary.shipNumber = Range("ae5").Value
    order.header.summary.shipDate = Range("ae9").Value
    order.header.summary.customerNumber = Range("ac47").Value
    order.header.summary.poNumber = Range("ac49").Value
    order.header.summary.poDate = Range("ac51").Value
    order.header.summary.vmOrderNumber = Range("ac53").Value
    order.header.summary.qty = 0
    order.header.summary.netTotalWeight = 0
    order.header.summary.grossTotalWeight = 0
    
    order.header.billTo.name = Range("f30").Value
    order.header.billTo.address1 = Range("f32").Value
    order.header.billTo.address2 = Range("f34").Value
    order.header.billTo.address3 = Range("f36").Value
    order.header.billTo.address4 = Range("f38").Value
    order.header.billTo.address5 = Range("f40").Value

    order.header.shipTo.name = Range("v30").Value
    order.header.shipTo.address1 = Range("v32").Value
    order.header.shipTo.address2 = Range("v34").Value
    order.header.shipTo.address3 = Range("v36").Value
    order.header.shipTo.address4 = Range("v38").Value
    order.header.shipTo.address5 = Range("v40").Value
    
    'loop over the data rows and add 'salesOrderDataLine' to the 'order.body.data' collection
    
    'the starting cell for collecting data: maybe i should save as array first? **
    Dim r As Range
    Set r = Range("c60")
    
    'get the product data well need for the lookups
    Dim productData() As Variant
    productData = getProductData()
    
    'loop from row 60 to the end of the page
    Dim i As Integer
    For i = 0 To Cells.SpecialCells(xlLastCell).Row - 60
    
        If IsNumeric(r.Offset(i, 0).Value) And r.Offset(i, 0).Value > 999 And r.Offset(i, 25).Value > 0 Then
            
            'vba has a funny way of keeping the variable alive after i redefine it
            'if i declare x as a new instance of a class, and add it to a collection
            'inside the loop, all the elements of the collection are the last object added
            'weird
            'Dim x As New salesOrderDataLine
            
            order.body.data.Add New salesOrderDataLine
            
            'item number
            Dim item As String
            item = r.Offset(i, 0).Value
            order.body.data(order.body.data.count).item = item
            
            'master
            Dim master As Long
            If Len(item) = 6 Then
                master = Right(item, 3)
            ElseIf Len(item) = 4 Then
                master = Right(item, 4)
            Else
                master = 999
            End If
            order.body.data(order.body.data.count).master = master
                                      
            'Debug.Print x.item & ", " & x.master
                    
            'desc
            Dim desc As String
            desc = r.Offset(i, 4).Value
            order.body.data(order.body.data.count).desc = ALOOKUP2(master, productData, 2, 1, 0)
            
            'unit
            Dim unit As String
            unit = r.Offset(i, 14).Value
            order.body.data(order.body.data.count).unit = unit
            
            'Debug.Print x.desc & ", " & x.unit
            
            'ordered
            Dim ordered As Long
            If IsNumeric(r.Offset(i, 25).Value) Then
                ordered = r.Offset(i, 25).Value
            Else
                ordered = 0
            End If
            order.body.data(order.body.data.count).ordered = ordered
            
            'shipped
            Dim shipped As Long
            If IsNumeric(r.Offset(i, 17).Value) Then
                shipped = r.Offset(i, 17).Value
            Else
                shipped = 0
            End If
            order.body.data(order.body.data.count).shipped = shipped
            order.header.summary.qty = order.header.summary.qty + shipped

            'lot: only if something was shipped
            If IsNumeric(shipped) And shipped > 0 Then
                Dim str As String
                str = r.Offset(i + 2, 4).Value
                order.body.data(order.body.data.count).lots = parseStringIntoLots(str)
            Else
                Dim arr2(0 To 0) As Variant
                arr2(0) = "x"
                order.body.data(order.body.data.count).lots = arr2
            End If
            
            'net unit weight
            Dim netUnitWeight As Double
            netUnitWeight = ALOOKUP2(master, productData, 7, 1, 0)
            order.body.data(order.body.data.count).netUnitWeight = netUnitWeight
            
            'net weight
            Dim netWeight As Double
            netWeight = netUnitWeight * shipped
            order.body.data(order.body.data.count).netWeight = netWeight
            order.header.summary.netTotalWeight = order.header.summary.netTotalWeight + netWeight
            
            'gross weight
            Dim grossWeight As Double
            grossWeight = netWeight * 1.1
            order.body.data(order.body.data.count).grossWeight = grossWeight
            order.header.summary.grossTotalWeight = order.header.summary.grossTotalWeight + grossWeight
        
        End If
    
    Next i
    
    Set getPicklistDataActions = order
    
    
End Function


