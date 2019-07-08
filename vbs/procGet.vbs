Option Explicit

' variable declaration
Dim svcs
Dim procList
Dim proc
Dim msg

' Process List Get 
Set svcs = WScript.CreateObject("WbemScripting.SWbemLocator").ConnectServer
Set procList = svcs.ExecQuery("Select * From Win32_Process")

' Process name Get
For Each proc In procList
    msg = msg & proc.Description & vbCrLf
Next

' Output
MsgBox msg

Set svcs = Nothing
Set procList = Nothing
Set proc = Nothing
Set msg = Nothing