VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "salesOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public header As New salesOrderHeader

Public body As New salesOrderBody


'method
Public Sub printHeader()

    Debug.Print header.summary.shipNumber
    Debug.Print header.summary.shipDate
    
    Debug.Print header.summary.customerNumber
    Debug.Print header.summary.poNumber
    Debug.Print header.summary.poDate
    Debug.Print header.summary.vmOrderNumber
    Debug.Print header.summary.qty
    Debug.Print header.summary.netTotalWeight
    Debug.Print header.summary.grossTotalWeight
    
End Sub
