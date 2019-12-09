VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "salesOrderDataLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public item As String, desc As String, unit As String

Public master As Integer, ordered As Integer, shipped As Integer

Public lots As Variant

Public netUnitWeight As Double, netWeight As Double, grossWeight As Double


