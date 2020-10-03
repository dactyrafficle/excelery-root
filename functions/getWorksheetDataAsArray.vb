Option Explicit


'there should be some default to get data for the whole sheet
'and for the activebook

Function getWorksheetDataAsArray(wsName As String) As Variant

  Dim ws As Worksheet
  Set ws = ThisWorkbook.Sheets(wsName)

  Dim r As Range
  Set r = ws.Range("a1")

  Dim n_rows, n_cols As Integer
  n_rows = r.End(xlDown).Row
  n_cols = r.End(xlToRight).Column

  'use -1 bc of how offset works
  Dim arr As Variant
  arr = ws.Range(r, r.Offset(n_rows - 1, n_cols - 1).Address)

  getWorksheetDataAsArray = arr

End Function

Function getWorksheetDataColumnHeadersAsArray(wsName As String) As Variant

  Dim ws As Worksheet
  Set ws = ThisWorkbook.Sheets(wsName)

  Dim r As Range
  Set r = ws.Range("a1")

  Dim n_rows, n_cols As Integer
  n_rows = r.End(xlDown).Row
  n_cols = r.End(xlToRight).Column

  'use -1 bc of how offset works
  Dim arr As Variant
  arr = ws.Range(r, r.Offset(0, n_cols - 1).Address)

  getWorksheetDataColumnHeadersAsArray = arr

End Function