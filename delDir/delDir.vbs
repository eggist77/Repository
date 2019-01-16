

'iomode
Const ForReading = 1, ForWriting = 2, ForAppending = 8

Dim fso

txtFile = "delDir.log"
dir = "delDirTest"

Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.OpenTextFile(txtFile, ForAppending, True)

If fso.FolderExists(dir) Then
  fso.DeleteFolder(dir)
  f.WriteLine(dir & "削除しました")
End If

f.Close
