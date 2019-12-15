
'use this to fill in the worksheet
'the use tacocat code to make the json file
'and the rest is tacocat


Sub abc()

    Dim HostFolder As String
    HostFolder = Application.ActiveWorkbook.Path & "\Active"
    Debug.Print HostFolder

    Dim FileSystem As Object
    Set FileSystem = CreateObject("Scripting.FileSystemObject")
    'Call DoFolder(FileSystem.GetFolder(HostFolder))
    DoFolder FileSystem.GetFolder(HostFolder)
    
End Sub

'im confused by the syntax
Sub DoFolder(Folder)

    'loop over all the subfolders
    Dim SubFolder As Object
    For Each SubFolder In Folder.SubFolders
        DoFolder SubFolder
        Debug.Print SubFolder
    Next
    
    'or loop over all the files
    Dim File
    For Each File In Folder.Files
        ' Operate on each file
        'Debug.Print File
    Next
    
End Sub
