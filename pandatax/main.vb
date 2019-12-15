
'use this to fill in the worksheet
'the use tacocat code to make the json file
'and the rest is tacocat


Sub abc()

    Dim HostFolder As String
    HostFolder = Application.ActiveWorkbook.Path & "\Active"
    'Debug.Print HostFolder

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
        
        Debug.Print SubFolder.Name
        
        Dim File
        For Each File In SubFolder.Files
        
            Debug.Print File.Name
        
        Next
        
    Next

    
End Sub
