
'iomode
Const ForReading = 1, ForWriting = 2, ForAppending = 8

txtFile = "testlog.txt"
txtFile2 = "result.csv"

Set fso = WScript.CreateObject("Scripting.FileSystemObject")
Set f = fso.OpenTextFile(txtFile, ForReading)
Set f2 = fso.OpenTextFile(txtFile2, ForWriting, True)

Do Until f.AtEndOfStream
  line = f.ReadLine

  delimiter1 = instr(1,line,": %",vbTextCompare)
  if delimiter1>0 then
    i = delimiter1+2
    delimiter2 = instr(i,line,": ",vbTextCompare)
  End if

  if delimiter1>0 and delimiter2>delimiter1 then
    msg1 = Left(line,delimiter1-1)
    msg2 = Mid(line,i,delimiter2-i)
    msg3 = """" & Mid(line,delimiter2+2) & """"
    f2.WriteLine msg1 & "," & msg2 & "," & msg3
  End if
Loop

f.Close
f2.Close
