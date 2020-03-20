
'nLib Version = 1.2

Function dateFormat(ByVal format)

  ' format
  ' yyyymm
  ' yyyymmdd(default)
  ' yyyymmddhhmmss
  ' yyyymmdd_hhmmss
  ' yyyymmddhhmm
  ' yyyymmdd_hhmm

  format = LCase(format)

  Select Case format
  Case "yyyymm"
    dateFormat = Replace(Left(Now(),8), "/", "")
  Case "yyyymmdd"
    dateFormat = Replace(Left(Now(),10), "/", "")
  Case "yyyymmddhhmmss"
    dateFormat = Replace(Mid(Now(),12), ":", "")
    if Len(dateFormat) = 5 Then dateFormat = "0" & dateFormat
    dateFormat = Replace(Left(Now(),10), "/", "") & dateFormat
  Case "yyyymmdd_hhmmss"
    dateFormat = Replace(Mid(Now(),12), ":", "")
    if Len(dateFormat) = 5 Then dateFormat = "0" & dateFormat
    dateFormat = Replace(Left(Now(),10), "/", "") & "_" & dateFormat
  Case "yyyymmddhhmm"
    dateFormat = Replace(Mid(Now(),12), ":", "")
    if Len(dateFormat) = 5 Then dateFormat = "0" & dateFormat
    dateFormat = Replace(Left(Now(),10), "/", "") & Left(dateFormat,4)
  Case "yyyymmdd_hhmm"
    dateFormat = Replace(Mid(Now(),12), ":", "")
    if Len(dateFormat) = 5 Then dateFormat = "0" & dateFormat
    dateFormat = Replace(Left(Now(),10), "/", "") & "_" & Left(dateFormat,4)
  Case Else
    dateFormat = Replace(Left(Now(),10), "/", "")
  End Select
End Function

Function isTextOnlyInFolder(ByVal folderName)

    Dim fso
    Dim folder
    Dim iNotText

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderName)

    'File Check
    iNotText = 0
    For Each tmp In folder.Files

        WScript.Echo "file," & tmp.name & "," & tmp.type
        If tmp.type <> "テキスト ドキュメント" Then
            iNotText = iNotText + 1
        End If
    Next

    WScript.Echo "Folder Count: " & folder.SubFolders.Count & vbCrLf & _
                 "Folder Count: " & iNotText

    'Judgment
    isTextOnlyInFolder = False
    If folder.SubFolders.Count = 0 And iNotText = 0 Then
        isTextOnlyInFolder = True
    Else
        isTextOnlyInFolder = False
    End If
End Function

Function getParameter(ByVal txt, ByVal delimiter)

    buf = inStr(txt,delimiter)
    If buf > 0 then
        getParameter = trim(mid(txt,buf + 1,len(txt)))
    end If

End Function
