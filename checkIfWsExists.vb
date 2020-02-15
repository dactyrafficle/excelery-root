Option Explicit

'THIS IS PROBABLY THE BEST TO INCORPORATE INTO A LARGE PIECE OF CODE
Function checkIfWsExists(wsName As String) As Boolean

  checkIfWsExists = False
    
  Dim i As Long
  For i = 1 To Worksheets.Count
    If Worksheets(i).Name = wsName Then
      checkIfWsExists = True
      Exit Function
    End If
  Next i
 
End Function

'AS A STAND-ALONE SUB THIS WORKS WELL, BUT THE EXIT CONDITION WONT EXIT THE PARENT
'THAT COULD LEAD TO PROBLEMS
Sub checkIfWsExistsElseCreate(wsName As String)

  Dim i As Long
  For i = 1 To Worksheets.Count
    If Worksheets(i).Name = wsName Then
      MsgBox "The ws ws(" & wsName & ") already exists."
      Exit Sub
    End If
  Next i
    
  Dim ws As Worksheet
  Set ws = Worksheets.Add
  ws.Name = wsName
    
End Sub


'HERES A LONGER VERSION
Sub checkIfWsExistsElseCreate2(wsName As String)

  Dim wsExists As Boolean
  wsExists = False
    
  Dim i As Long
  For i = 1 To Worksheets.Count
    If Worksheets(i).Name = wsName Then
      wsExists = True
      Exit For
    End If
  Next i
    
  If wsExists Then
    MsgBox "The ws ws(" & wsName & ") already exists."
    Exit Sub
  End If

  Dim ws As Worksheet
  Set ws = Worksheets.Add
  ws.Name = wsName
    
End Sub