Set wmiServ=GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")

Set fso = CreateObject("Scripting.FileSystemObject")
toolDir = fso.getParentFolderName(WScript.ScriptFullName)

Set f = fso.CreateTextFile(toolDir & "\result.txt",True)
 
For Each tmpClass in wmiServ.SubclassesOf()
    f.WriteLine tmpClass.Path_.Class
Next
