

toZip_Lpls("C:\nBox\share\03_PC\GitHub\vbs\toZip_Lpls\test")

Function toZip_Lpls(target)

Dim wsh

Set wsh = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

libPath = "C:\nBox\share\03_PC\GitHub\vbs\toZip_Lpls\lib"
Lpls = libPath & "\Lhaplus.exe"

If fso.FileExists(Lpls) then
  If fso.FileExists(libPath & "/Extract1.dll") then

    targetPath = fso.getParentFolderName(target)
    command = Lpls & " /c:zip /o:" & targetPath & " " & target
    wsh.Run command, 0, true
  Else
    msgbox "Lhaplus not found!"
  End If
Else
  msgbox "Lhaplus not found!"
End If

End Function
