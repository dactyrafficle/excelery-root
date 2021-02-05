Option Explicit

Sub abc()

 Dim oApp As Outlook.Application
 Dim oNs As Outlook.NameSpace
 Dim oFolder As Outlook.MAPIFolder

 Set oApp = New Outlook.Application
 Set oNs = oApp.GetNamespace("MAPI")

 For Each oFolder In oNs.Folders
  processFolder oFolder
 Next

End Sub

 

Private Sub processFolder(ByVal oParent As Outlook.MAPIFolder, Optional ByVal path As String = "")

 path = path & "\" & oParent.Name
 Debug.Print path

 Dim oFolder As Outlook.MAPIFolder
 Dim oMail As Outlook.MailItem

 'For Each oMail In oParent.Items
 'Next

 If (oParent.Folders.Count > 0) Then
  For Each oFolder In oParent.Folders
   processFolder oFolder, path
  Next
 End If

End Sub
