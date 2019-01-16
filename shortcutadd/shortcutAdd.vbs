
Option Explicit

'variable declaration
Dim crDir, f, line, ary, fName
Dim fso, wsh

Set fso = CreateObject("Scripting.FileSystemObject")
Set wsh = CreateObject("WScript.Shell")

crDir = fso.getParentFolderName(WScript.ScriptFullName)
Set f = fso.OpenTextFile(crDir & "\list.txt", 1)

Do Until f.AtEndOfStream
  line = f.ReadLine
  ary = Split(line, ",")
  Call shortCutAdd(ary(0),ary(1))
Loop

f.Close


Sub shortCutAdd(scname,exePath)

Dim shortCutFile

  fName = crDir & "\" & scname & ".lnk"

  Set shortCutFile = wsh.CreateShortcut(fName)
  shortCutFile.TargetPath = exePath
  shortCutFile.Save
End Sub
