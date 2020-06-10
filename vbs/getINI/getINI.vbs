

file = "C:\Users\eggis\GDr\03_PC\GitHub\vbs\getINI\test.ini"

msgbox getINI("Section1","key1",file)


Function getINI(sectionName,keyName,fileName)

	Set iniDic = readINI(fileName)
	getINI = iniDic.Item(sectionName).Item(keyName)
End Function

Function readINI(fileName)

	Dim fso
	Dim sectionDic	'section Dictionary'
	Dim dic  		'Dictionary
	Dim f  			'file
	Dim line

	Dim sectionName : sectionName = ""
	Dim arrValue

	Set fso = CreateObject("Scripting.FileSystemObject")
	Set dic = CreateObject("Scripting.Dictionary")

	If fso.FileExists(fileName) Then
		Set f = fso.OpenTextFile(fileName)

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
				arrValue = Split(line,"=")
				dic(sectionName).Add Trim(arrValue(0)), Trim(arrValue(1))

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
