
'Option Explicit

'variable declaration
Dim fso, wsh
Dim crDir, f, line, ary, fName

Call mail

Sub mail()

    Dim scList : scList = "sclist.csv"

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set wsh = CreateObject("WScript.Shell")

    crDir = fso.getParentFolderName(WScript.ScriptFullName)

    'File Check
    If fso.FileExists(crDir & "\" & scList) then

        Set f = fso.OpenTextFile(crDir & "\" & scList, 1)

        Do Until f.AtEndOfStream
          line = f.ReadLine
          If line = "Name,TargetPath" Then line = f.ReadLine 'Header Skip'
          ary = Split(line, ",")
          Call shortCutAdd(ary(0),ary(1))
        Loop
        f.Close
    Else
        msgbox scList & " file not found"
    End If
End Sub

Sub shortCutAdd(scname,exePath)

Dim shortCutFile

  fName = crDir & "\" & scname & ".lnk"

  Set shortCutFile = wsh.CreateShortcut(fName)
  shortCutFile.TargetPath = exePath
  shortCutFile.Save
End Sub
