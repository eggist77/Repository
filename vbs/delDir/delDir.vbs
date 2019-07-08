

main()

Sub main()

'iomode
Const ForReading = 1, ForWriting = 2, ForAppending = 8

Dim fso

txtFile = "delDir.log"
dir = "delDirTest"
daysBefore = 1
delFlag = False

Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.OpenTextFile(txtFile, ForAppending, True)
toolDir = fso.getParentFolderName(WScript.ScriptFullName)

Set items = fso.GetFolder(toolDir)

For Each item In items.SubFolders

  ' 作成日付チェック
  WScript.Echo item.Name & " " & item.DateCreated
  WScript.Echo DateDiff("d", now(), item.DateCreated)

  If DateDiff("d", now(), item.DateCreated) = 0 then

    ' フォルダの中身をチェック
    Set items2 = fso.GetFolder(toolDir & "\" & item.Name)

    ' FolderCheck
    folderCnt = 0
    For Each item2 In items2.SubFolders
      folderCnt = folderCnt + 1
    Next

    if folderCnt = 0 then
      ' FileCheck
      fileCnt = 0
      For Each item2 In items2.Files
        ' textCheck
        msgbox item2.name & " " & item2.type
      Next
    End If

    msgbox "folderCnt:" & folderCnt & vbCrlf & "fileCnt:" & fileCnt

    ' Folder Delete
    msgbox item.Name & "    " & cnt
    delFlag = False

  End If
Next

exit Sub

If fso.FolderExists(dir) Then
  fso.DeleteFolder(dir)
  f.WriteLine(dir & "削除しました")
End If

f.Close

End Sub
