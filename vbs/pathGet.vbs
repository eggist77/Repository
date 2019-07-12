


Set fso = CreateObject("Scripting.FileSystemObject")

' Get ScriptFile folder
msgbox fso.getParentFolderName(WScript.ScriptFullName)
