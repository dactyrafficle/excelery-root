Option Explicit

Sub validate_info1()

  If LCase(ActiveSheet.Name) <> "info1" Then
    MsgBox "This isn't ws(" & Chr(34) & "Info1" & Chr(34) & "). Don't run this here"
    Exit Sub
  End If
    
  Dim ws As Worksheet
  Set ws = Worksheets("info1")

  Dim r As Range
  Set r = ws.Range("m3")
    
  r.Offset(0, 0).Value = "ws_name"
  r.Offset(0, 1).Value = "criteria_field"
  r.Offset(0, 2).Value = "sum_field"
  r.Offset(0, 3).Value = "TOTAL CASES:"
  r.Offset(0, 4).Value = "TOTAL VALUE:"

  Dim i As Long
  For i = 1 To Worksheets.Count

    r.Offset(i, 0).Value = Worksheets(i).Name
    r.Offset(i, 1).Value = "'" & Chr(39) & Worksheets(i).Name & "'!j:j"
    r.Offset(i, 2).Value = "'" & Chr(39) & Worksheets(i).Name & "'!k:k"
    r.Offset(i, 3).Formula = "=sumifs(indirect(" & r.Offset(i, 2).Address(False, False) & "),indirect(" & r.Offset(i, 1).Address(False, False) & ")," & Chr(34) & "TOTAL CASES:" & Chr(34) & ")"
    r.Offset(i, 4).Formula = "=sumifs(indirect(" & r.Offset(i, 2).Address(False, False) & "),indirect(" & r.Offset(i, 1).Address(False, False) & ")," & Chr(34) & "TOTAL VALUE:" & Chr(34) & ")"


  Next i

End Sub
