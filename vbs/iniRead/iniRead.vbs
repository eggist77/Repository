'Ini�t�@�C�����擾
Function GetIniFileInfo(filepath)
  Dim objFSO  'FileSystemObject
  Dim objInfo  'ini�t�@�C���pDictionary
  Dim objFile  'ini�t�@�C��
  Dim strLine  '�f�[�^�ǂݍ��ݗp
  Dim objCat  '�J�e�S���pDictionary
  Dim strCat  '�J�e�S����
  Dim arrValue'�L�[�ƒl�̔z��
  
  Set objFSO = CreateObject("Scripting.FileSystemObject")
  Set objInfo = CreateObject("Scripting.Dictionary")
  
  'ini�t�@�C�������݂���΃f�[�^�擾
  If objFSO.FileExists(filepath) Then
    Set objFile = objFSO.OpenTextFile(filepath)
    
    Do Until objFile.AtEndOfStream
      strLine = Trim(objFile.ReadLine)
      
      '�J�e�S����\���s��
      If Left(strLine,1) = "[" And Right(strLine,1) = "]" Then
        'Category
        strCat = Mid(strLine,2,Len(strLine) - 2)
        If Not objInfo.Exists(strCat) Then
          Set objCat = CreateObject("Scripting.Dictionary")
          objInfo.Add strCat,objCat
        End If
      ElseIf Instr(strLine,"=") > 1 And strCat <> "" Then
        'Key & Value
        arrValue = Split(strLine,"=")
        objInfo(strCat).Add Trim(arrValue(0)), _
              Trim(Mid(Join(arrValue,"="),Len(arrValue(0)) + 2))
      End If
    Loop
    
    objFile.Close
  End If
  
  Set GetIniFileInfo = objInfo
  
  Set objFile = Nothing
  Set objCat = Nothing
  Set objInfo = Nothing
  Set objFSO = Nothing
End Function