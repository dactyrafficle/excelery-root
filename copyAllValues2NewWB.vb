Option Explicit

Sub copyAllValues2NewWB()

  Dim wb1 As Workbook
  Set wb1 = ActiveWorkbook

  'the reason i didnt set the number of sheets to the source wb is bc i dont want the names flying around
  Application.SheetsInNewWorkbook = 1
  Dim wb2 As Workbook
  Set wb2 = Workbooks.Add

  Dim i As Long
  For i = 1 To wb1.Sheets.Count

    Dim arr As Variant
    arr = wb1.Sheets(i).UsedRange.Value
    
    wb2.Sheets(i).name = wb1.Sheets(i).name
    wb2.Sheets(i).Range("a1").Resize(UBound(arr, 1), UBound(arr, 2)).Value = arr

    If i <> wb1.Sheets.Count Then
      wb2.Sheets.Add after:=wb2.Sheets(wb2.Sheets.Count)
    End If

  Next i

End Sub