
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

        If fso.FolderExists(crDir & "\output") = false then
            fso.CreateFolder(crDir & "\output")
        End If

        'WScript.Sleep 2000
        Set f = fso.OpenTextFile(crDir & "\" & scList, 1)

        Do Until f.AtEndOfStream
          line = f.ReadLine
          'Header and comment Skip'
          If line = "Name,TargetPath" or Left(line,1) = "'" Then line = f.ReadLine

          If instr(line,",") > 0 then
            ary = Split(line, ",")
            Call shortCutAdd(ary(0),ary(1))
          End If
        Loop
        f.Close
    Else
        msgbox scList & " file not found"
    End If
End Sub

Sub shortCutAdd(scname,exePath)

Dim shortCutFile

  fName = crDir & "\output\" & scname & ".lnk"

  Set shortCutFile = wsh.CreateShortcut(fName)
  shortCutFile.TargetPath = exePath
  shortCutFile.Save
End Sub
