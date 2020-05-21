
Set fso = CreateObject("Scripting.FileSystemObject")

folderPath = "."

getFileList fso.getFolder(folderPath)

sub getFileList(ByVal folder,ByVal level)

    for each subFolder in folder.subfolders
        WScript.Echo "d," & subFolder.name & "," & subFolder.size & "," & subFolder.ParentFolder

        'folder List Output'
    next

    for each f in folder.files
        WScript.Echo "-," & f.name & "," & f.size & "," & f.ParentFolder
    next

    for 1 to level
        'もしフォルダリストがあったら・・・'
        If XXX then
            for each subFolder in folder.subfolders
                WScript.Echo "d," & subFolder.name & "," & subFolder.size & "," & subFolder.ParentFolder

                'folder List Output'
            next

            for each f in folder.files
                WScript.Echo "-," & f.name & "," & f.size & "," & f.ParentFolder
            next
        End If
    Next 
End sub
