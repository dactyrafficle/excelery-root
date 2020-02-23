Option Explicit

Sub cleanMasterFromColumnB()

  Dim i As Long
  For i = 1 To Worksheets.Count
    
    Dim isMasterColumn As Boolean, y_start As Long
    isMasterColumn = False
    
    Dim y As Long
    For y = 1 To Worksheets(i).Cells.SpecialCells(xlLastCell).Row
      If Worksheets(i).Range("b" & y).Value = "CARTON CODE" Then    
        isMasterColumn = True
        y_start = y + 1
        Debug.Print Worksheets(i).Name & ": " & Worksheets(i).Range("b" & y).Address
        With Worksheets(i).Range("b" & y)
          .Style = "Input"
          .Style = "Good"
          .HorizontalAlignment = xlCenter
          .VerticalAlignment = xlCenter
          .WrapText = True
        End With        
        Exit For
      End If
    Next y
        
    If isMasterColumn Then
      For y = y_start To Worksheets(i).Cells.SpecialCells(xlLastCell).Row
        If Len(Trim(Worksheets(i).Range("b" & y).Value)) > 3 Then
          Worksheets(i).Range("b" & y).Value = Left(Trim(Worksheets(i).Range("b" & y).Value), 3)
        End If
      Next y
    End If

  Next i

End Sub
