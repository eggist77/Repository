

main()

Sub main()

'iomode
Const ForReading = 1, ForWriting = 2, ForAppending = 8

Dim fso

txtFile = "delDir.log"
dir = "delDirTest"
daysBefore = 7
delFlag = False

Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.OpenTextFile(txtFile, ForAppending, True)
toolDir = fso.getParentFolderName(WScript.ScriptFullName)

Set items = fso.GetFolder(toolDir)

For Each item In items.SubFolders

  'If DateDiff("d", now(), item.DateCreated) = 0 then
  If DateDiff("s", item.DateCreated, now()) >= daysBefore then

    f.WriteLine now() & "," & item.name & "," & item.DateCreated

    '
    Set items2 = fso.GetFolder(toolDir & "\" & item.Name)

    ' FolderCheck
    If items2.SubFolders.Count = 0 Then

      ' FileCheck
      fileCnt = 0
      For Each item2 In items2.Files
        ' textCheck
        msgbox item2.name & " " & item2.type
      Next

      If fileCnt = 0 Then
      End If
    Else
      f.WriteLine now() & "," & item.name & ",Folder Find"

    End If

    ' Folder Delete
    msgbox item.Name & "    " & cnt
    delFlag = False

  End If
Next

exit Sub

If fso.FolderExists(dir) Then
  fso.DeleteFolder(dir)
  f.WriteLine(dir & "")
End If

f.Close

End Sub
