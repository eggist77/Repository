'Ini�t�@�C�����擾

res = getINI("test","test","test.ini") 

Function getINI(iniSection,iniKey,fileName)

'�쐬���B�܂������Ȃ�'

msgbox "test"

	Dim fso
	Dim dic  		'ini�t�@�C���pDictionary
	Dim f  			'ini�t�@�C��
	Dim line  		'�f�[�^�ǂݍ��ݗp
	Dim sectionDic  	'�Z�N�V�����pDictionary
	Dim sectionName  	'�Z�N�V������
	Dim arrValue	'�L�[�ƒl�̔z��

	iniSection = UCase(iniSection)
	iniKey = UCase(iniKey)

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

			ElseIf Left(line,1) = ";" Then

			'Parameter
			ElseIf Instr(line,"=") > 1 And sectionName <> "" Then

				'Key & Value
				arrValue = Split(line,"=")
				dic(sectionName).Add Trim(arrValue(0)), Trim(Mid(Join(arrValue,"="), Len(arrValue(0)) + 2))
			End If
		Loop

		f.Close

		'Dictionary�I�u�W�F�N�g�̗v�f�̎Q��
   		Dim str
   		For Each Var In dic
	   		str = str & Var & " : " & dic.Item(Var) & vbCrLf
   		Next

   		MsgBox str, vbInformation

	End If

	Set f = Nothing
	Set sectionDic = Nothing
	Set dic = Nothing
	Set fso = Nothing

End Function
