

main()

sub main()

    Dim fso
    Dim folder

    folderName = "C:\Users\eggis\GDr\03_PC\GitHub\vbs"

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderName)

    For Each tmp In folder.SubFolders

        WScript.Echo "folder," & tmp.name & "," & tmp.size & "bytes"
    Next

    For Each tmp In folder.Files

        WScript.Echo "file," & tmp.name & "," & tmp.size& "bytes"
    Next

End Sub
