
'Write to a text file

'iomode
Const ForReading = 1, ForWriting = 2, ForAppending = 8

txtFile = "txtfile.txt"

Set fso = WScript.CreateObject("Scripting.FileSystemObject")
Set f = fso.OpenTextFile(txtFile, ForAppending, True)

' header
If f.line = 1 Then
  f.WriteLine "title1,title2,title3"
End If

f.WriteLine "data1,data2,data3"

f.Close
