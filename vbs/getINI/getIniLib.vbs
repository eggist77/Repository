
Function getINI(sectionName,keyName,iniFile)

    Set iniDic = readINI(iniFile)
    getINI = iniDic.Item(sectionName).Item(keyName)
End Function

Function readINI(iniFile)

    Dim fso         'file system object
    Dim sectionDic  'section Dictionary
    Dim dic         'Dictionary
    Dim f           'file
    Dim line
    Dim sectionName : sectionName = ""
    Dim a           'array

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set dic = CreateObject("Scripting.Dictionary")

    If fso.FileExists(iniFile) Then
        Set f = fso.OpenTextFile(iniFile)

        Do Until f.AtEndOfStream
            line = Trim(f.ReadLine)

            'Section
            If Left(line,1) = "[" And Right(line,1) = "]" Then
                sectionName = Mid(line, 2, Len(line) - 2)

                If Not dic.Exists(sectionName) Then
                  Set sectionDic = CreateObject("Scripting.Dictionary")
                  dic.Add sectionName, sectionDic
                End If

            'Parameter
            ElseIf Instr(line,"=") > 1 And sectionName <> "" Then

                'Key & Value
                a = Split(line,"=")
                dic(sectionName).Add Trim(a(0)), Trim(a(1))

            'comment'
            ElseIf Left(line,1) = ";" Then
            End If
        Loop

        Set readINI = dic
        f.Close
    End If

    Set f = Nothing
    Set sectionDic = Nothing
    Set dic = Nothing
    Set fso = Nothing
End Function
