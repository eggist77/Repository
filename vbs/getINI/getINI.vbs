'Ini�t�@�C�����擾
Function GetIniFileInfo(filepath)

Dim fso
Dim dic  'ini�t�@�C���pDictionary
Dim f  'ini�t�@�C��
Dim line  '�f�[�^�ǂݍ��ݗp
Dim catDic  '�J�e�S���pDictionary
Dim secName  '�J�e�S����
Dim arrValue'�L�[�ƒl�̔z��

Set fso = CreateObject("Scripting.FileSystemObject")
Set dic = CreateObject("Scripting.Dictionary")

If fso.FileExists(filepath) Then
	Set f = fso.OpenTextFile(filepath)

	Do Until f.AtEndOfStream
		line = Trim(f.ReadLine)

		'Section
		If Left(line,1) = "[" And Right(line,1) = "]" Then

			'Section
			secName = Mid(line, 2, Len(line) - 2)

			If Not dic.Exists(secName) Then

			  Set catDic = CreateObject("Scripting.Dictionary")
			  dic.Add secName, catDic
			End If
		
		'Parameter
		ElseIf Instr(line,"=") > 1 And secName <> "" Then

			'Key & Value
			arrValue = Split(line,"=")
			dic(secName).Add Trim(arrValue(0)), Trim(Mid(Join(arrValue,"="), Len(arrValue(0)) + 2))
		End If
	Loop

	f.Close
End If

Set GetIniFileInfo = dic

Set f = Nothing
Set catDic = Nothing
Set dic = Nothing
Set fso = Nothing
End Function