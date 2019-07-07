

main()

Sub main()

'iomode
Const ForReading = 1, ForWriting = 2, ForAppending = 8

Dim fso

txtFile = "delDir.log"
dir = "delDirTest"
daysBefore = 1

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
    cnt = 0
    For Each item2 In items2.SubFolders
      cnt = cnt + 1
    Next

    if cnt = 0 then

      ' FileCheck
      cnt = 0
      For Each item2 In items2.Files
        msgbox item2.Name & "    " & item2.type
      Next
    End If


    msgbox item.Name & "    " & cnt



  End If
Next

exit Sub

If fso.FolderExists(dir) Then
  fso.DeleteFolder(dir)
  f.WriteLine(dir & "削除しました")
End If

f.Close

End Sub
