
'Option Explicit

'variable declaration
Dim fso, wsh
Dim crDir, f, line, ary, fName

Call mail


Sub mail()

    Dim scList : scList = "sclist.csv"
    Dim excel
    Dim folder

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set wsh = CreateObject("WScript.Shell")
    Set excel = CreateObject("Excel.Application")

    crDir = fso.getParentFolderName(WScript.ScriptFullName)

    'DefaultFilePath Change'
    excel.DefaultFilePath = crDir
    Set excel = Nothing
    Set excel = CreateObject("Excel.Application")

    dlg = excel.FileDialog(msoFileDialogFolderPicker)
    dlg.Show

    End If


    Exit Sub

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

Sub shortCutList(scname,exePath)

Dim shortCutFile

  fName = crDir & "\" & scname & ".lnk"

  Set shortCutFile = wsh.CreateShortcut(fName)
  shortCutFile.TargetPath = exePath
  shortCutFile.Save
End Sub
