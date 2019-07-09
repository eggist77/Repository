

delDir()

Function delDir()

'iomode
Const ForReading = 1, ForWriting = 2, ForAppending = 8

Dim fso, wsh
Dim log

targetPath = ""
txtFile = "delDir.log"
daysBefore = 7

Set fso = CreateObject("Scripting.FileSystemObject")
Set wsh = CreateObject("WScript.Shell")
Set log = fso.OpenTextFile(txtFile, ForAppending, True)

programFilesPath = wsh.ExpandEnvironmentStrings("%ProgramFiles%")

msgbox programFilesPath

Private Const WindowsFolder = 0
Private Const SystemFolder = 1
Private Const TemporaryFolder = 2
 
MsgBox CreateObject("Scripting.FileSystemObject").GetSpecialFolder(TemporaryFolder).Path


' targetPath Check
'If objFso.FolderExists(targetPath) Then

'End If



Exit Function

Set items = fso.GetFolder(targetPath)

For Each item In items.SubFolders

  'If DateDiff("d", now(), item.DateCreated) = 0 then
  If DateDiff("s", item.DateCreated, now()) >= daysBefore then

    ' FolderCheck
    Set items2 = fso.GetFolder(toolDir & "\" & item.Name)

    ' FolderCheck No Folder
    If items2.SubFolders.Count = 0 Then

      ' FileCheck
      fileCnt = 0
      For Each item2 In items2.Files
        If item2.type <> "テキスト ドキュメント" Then
          fileCnt = fileCnt + 1
        End If
      Next

      If fileCnt = 0 Then
        log.WriteLine now() & "," & item.name & ",Folder Delete"
        item.delete
      Else
        log.WriteLine now() & "," & item.name & ",Not TextFile Find!"
      End If
    Else
      log.WriteLine now() & "," & item.name & ",Folder Find!"
    End If
  End If
Next

log.Close

End Function
