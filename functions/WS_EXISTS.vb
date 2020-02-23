Option Explicit

'maybe i should specify the workbook. because i think it will default to the activebook? or ThisWorkbook
Function WS_EXISTS(WS_NAME As String) As Boolean

  WS_EXISTS = False
    
  Dim i As Long
  For i = 1 To Worksheets.Count
    If Worksheets(i).Name = WS_NAME Then
      WS_EXISTS = True
      Exit Function
    End If
  Next i
 
End Function