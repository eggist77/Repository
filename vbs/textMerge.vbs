Option Explicit
Const OutFileName = "C:\work\Out\rusult.txt"
Const mergingTextFolder = "C:\work"
Const ForReading = 1, ForWriting = 2

Dim fso, f, folder, fileList, InFile, OutFile

Set fso = WScript.CreateObject("Scripting.FileSystemObject")
Set folder = fso.GetFolder(mergingTextFolder)

Set fileList = folder.Files

Set OutFile = fso.OpenTextFile(OutFileName, ForWriting, true)

For Each f in fileList

  If LCase(fso.GetExtensionName(f))="txt" Then

    Set InFile = fso.OpenTextFile(f, ForReading)
    OutFile.Write(InFile.ReadAll())
    OutFile.Write(vbCrLf)
    InFile.Close()
  End If
Next

OutFile.Close()
WScript.Echo "text Meage Complete!"
