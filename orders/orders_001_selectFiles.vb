Option Explicit

Sub openDialogBoxAndSelectFiles()

  Dim wb1 As Workbook
  Set wb1 = ActiveWorkbook

  With Application.FileDialog(msoFileDialogFilePicker)

    .AllowMultiSelect = True
    .InitialFileName = ThisWorkbook.path & "\"
    .Title = "Paddington Bear Selection Window"
    .ButtonName = "Omlette"
    
    .Filters.Clear
    .Filters.Add "All Files", "*.*"

    If .Show = True Then
    
      Dim file As Variant
      For Each file In .SelectedItems
        
      Dim wb2 As Workbook
      Set wb2 = Workbooks.Open(Filename:=file, ReadOnly:=True)
        
        Dim i As Long
        For i = 1 To wb2.Sheets.Count
          wb2.Sheets(i).Copy before:=wb1.Sheets(1)
        Next i
        
        wb2.Close
      Next
      
    End If
    
  End With
  
End Sub
