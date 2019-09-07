
'nLib Version = 1.1

Function dateFormat(format)

  ' format
  ' yyyymmdd
  ' yyyymmddhhmmss
  ' yyyymmddhhmm

  format = LCase(format)

  Select Case format
  Case "yyyymmdd"
    dateFormat = Replace(Left(Now(),10), "/", "")
  Case "yyyymmddhhmmss"
    dateFormat = Replace(Replace(Replace(Now(), "/", ""), ":", ""), " ", "")
  Case "yyyymmddhhmm"
    dateFormat = Replace(Replace(Replace(Left(Now(),16), "/", ""), ":", ""), " ", "")
  Case Else
    dateFormat = Replace(Left(Now(),10), "/", "")
  End Select
End Function

Function isTextOnlyInFolder(ByVal folderName)

    Dim fso
    Dim folder
    Dim iFolder, iNotText

    Set fso = CreateObject("Scripting.FileSystemObject")

    Set folder = fso.GetFolder(folderName)

    'SubFolder Check
    iFolder = 0
    For Each tmp In folder.SubFolders

        WScript.Echo "folder," & tmp.name
        iFolder = iFolder + 1
    Next

    'File Check
    iNotText = 0
    For Each tmp In folder.Files

        WScript.Echo "file," & tmp.name & "," & tmp.type
        If tmp.type <> "テキスト ドキュメント" Then
            iNotText = iNotText + 1
        End If
    Next

    WScript.Echo "Folder Count: " & iFolder & vbCrLf & "Folder Count: " & iNotText

    'Judgment
    isTextOnlyInFolder = False
    If iFolder = 0 And iNotText = 0 Then
        isTextOnlyInFolder = True
    Else
        isTextOnlyInFolder = False
    End If
End Function
