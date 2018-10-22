'Read a text file

'iomode
Const ForReading = 1, ForWriting = 2, ForAppending = 8

txtFile = "txtfile.txt"

Set fso = WScript.CreateObject("Scripting.FileSystemObject")
Set f = fso.OpenTextFile(txtFile, ForReading)

Do Until f.AtEndOfStream
  line = f.ReadLine
  msgbox line
Loop

f.Close
