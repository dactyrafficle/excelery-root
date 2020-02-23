Option Explicit

Sub createMasterDemand()

  'check if ws('master') exists
  Dim checkIfWsExists As Boolean
  checkIfWsExists = False

  Dim i As Long
  For i = 1 To Worksheets.Count
  If Worksheets(i).Name = "master" Then
    checkIfWsExists = True
    Exit For
  End If
  Next i

  If checkIfWsExists Then
    MsgBox "ws(master) already exists - try again"
    Exit Sub
  End If

  'if not exist, create it
  Dim ws As Worksheet
  Set ws = Worksheets.Add(Before:=Worksheets(1))
  ws.Name = "master"

  Dim n As Long
  n = 0

  'look at each ws except the first one - the one we just made
  For i = 2 To Worksheets.Count

    'cap the number of rows well look at
    Dim y_max As Long
    If Worksheets(i).Cells.SpecialCells(xlLastCell).Row > 120 Then
      y_max = 120
    Else
      y_max = Worksheets(i).Cells.SpecialCells(xlLastCell).Row
    End If
    
    'check column B for "CARTON CODE"
    'check column K for "quantity"

    Dim y As Long
    For y = 1 To y_max
    
      Dim qty As Variant
      qty = Sheets(i).Range("k" & y).Value

      Dim master As String
      master = Sheets(i).Range("b" & y).Value
        
      If IsNumeric(qty) Then
        If qty > 0 Then
          If Int(qty) / qty = 1 Then  
            If Trim(master & vbNullString) <> vbNullString Then 
              n = n + 1
              Sheets(1).Range("a" & n).Value = Sheets(i).Name
              Sheets(1).Range("b" & n).Formula = "=VLOOKUP(A1, Info1!B:E, 4, 0)"
              Sheets(1).Range("c" & n).Value = "year"
              Sheets(1).Range("d" & n).Value = "period"
              Sheets(1).Range("e" & n).Value = "xmas"
              Sheets(1).Range("f" & n).Value = master
              Sheets(1).Range("g" & n).Value = qty
              Sheets(1).Range("h" & n).Formula = Sheets(i).Range("j" & y).Value * qty     
            End If
          End If
        End If
      End If
    Next y
  Next i

End Sub