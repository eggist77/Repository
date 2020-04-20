
Set fso = CreateObject("Scripting.FileSystemObject")

folderPath = "."

getFileList fso.getFolder(folderPath)

sub getFileList(ByVal folder)

    for each subFolder in folder.subfolders
        WScript.Echo "d," & subFolder.name & "," & subFolder.size & "," & subFolder.ParentFolder
    next

    for each f in folder.files
        WScript.Echo "-," & f.name & "," & f.size & "," & f.ParentFolder
    next
End sub
