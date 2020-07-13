
'nLib Version = 1.4

'- List -
'dateFormat
'isTextOnlyInFolder
'getParameter
'getFilePathDlgIE
'getFilePathDlgExcel
'dec2bin

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

Function getFilePathDlgIE()
	Dim ie
	Set ie = WScript.CreateObject("InternetExplorer.Application")
	ie.Navigate "about:blank"
	Do While ie.Busy = True And ie.ReadyState <> 4 'READYSTATE_COMPLETE = 4
		WScript.Sleep 100
	Loop

	ie.document.write "<html><body><input type='file' id='selectFileDialog'></body></html>"
	ie.document.getElementById("selectFileDialog").click
	getFilePathDlgIE = ie.document.getElementById("selectFileDialog").Value

	ie.Quit
	Set objIE = Nothing
End Function

Function getFilePathDlgExcel()
	Dim excel
	Set excel = CreateObject("Excel.Application")
	buf = excel.GetOpenFilename("Text File,*.txt,All,*.*",1,"ファイルを選択して下さい","開く",false)
	If buf <> False Then
		getFilePathDlgExcel = buf
	Else
	    WScript.Quit
	End If
End Function

Function dec2bin(ByVal target)

    Dim amari()
    Dim i

    i = 0
    Do While target <> 0
        ReDim Preserve amari(i)
        amari(i) = target mod 2
        target = target \ 2
        i = i + 1
    Loop

    'list reverse
    i = 0
    For i = UBound(amari) To LBound(amari) Step -1
        buf = buf & amari(i)
    Next

    dec2bin = buf

End Function
