'use this to fill in the worksheet
'the use tacocat code to make the json file
'and the rest is tacocat


Sub abc()

    Dim HostFolder As String
    HostFolder = Application.ActiveWorkbook.Path & "\Active"
    
    Dim FileSystem As Object
    Set FileSystem = CreateObject("Scripting.FileSystemObject")
    
    'Call DoFolder(FileSystem.GetFolder(HostFolder))
    DoFolder FileSystem.GetFolder(HostFolder)
    
End Sub

'im confused by the syntax
Sub DoFolder(Folder)

    Dim col As New Collection

    'loop over all the subfolders
    Dim SubFolder As Object
    For Each SubFolder In Folder.SubFolders
        DoFolder SubFolder
        
        col.Add New Collection
        col(col.Count).Add SubFolder.Name

        Dim File
        For Each File In SubFolder.Files
            col(col.Count).Add File.Name
        Next
    Next
    
    Dim i As Long, j As Long
    For i = 1 To col.Count
        For j = 1 To col(i).Count
            Debug.Print col(i)(j)
        Next j
    Next i

End Sub
