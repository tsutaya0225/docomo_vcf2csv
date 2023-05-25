Option Explicit

'FileSystem�I�u�W�F�N�g���쐬����
Dim objFS : Set objFS=CreateObject("Scripting.FileSystemObject")

'Shell�I�u�W�F�N�g���쐬����
Dim objShell : Set objShell=CreateObject("WScript.Shell")

'�������Ȃ���΂��̂܂܏I������
If WScript.Arguments.Count = 0 Then
	MsgBox "�t�@�C�����h���b�v���Ă��������B", vbOKOnly+vbInformation, "CSV to VCF"
	WScript.Quit
End If

Dim inputFile, outputFile

'�����������̏ꍇ�͌J��Ԃ����s����
For Each inputFile In WScript.Arguments
	'�����̐�΃p�X���擾
	inputFile=LCase(objFS.getAbsolutePathName(inputFile))
	'���̓t�@�C���̊g���q���m�F
	If objFS.GetExtensionName(inputFile) <> "csv" Then
		MsgBox "CSV�t�@�C���ł͂���܂���B" & vbCrLf & "�I�����܂��B", vbOKOnly+vbCritical, "CSV to VCF"
		Wscript.Quit
	End If

	'�o�̓t�@�C������ݒ�
	outputFile=objFS.getParentFolderName(inputFile) & "\" & objFS.GetBaseName(inputFile) & ".vcf"
'	�o�̓t�@�C�����̊m�F
'	MsgBox "outputFile�� " & outputFile & " �ł�", vbOKOnly+vbInformation, "CSV to VCF"
	'�o�̓t�@�C�����쐬
	If objFS.FileExists(outputFile) Then
		if(MsgBox(outputFile & "�͊��ɑ��݂��܂��B" & vbCrLf & "�㏑�����܂����H", +vbExclamation, "CSV to VCF"))=vbCancel Then
			MsgBox "�I�����܂�"
			WScript.Quit
		End If
	End If

	'�e�L�X�g���擾����
	Dim inputText : set inputText= objFS.OpenTextFile(inputFile, 1)
	Dim outputText : Set outputText = objFS.OpenTextFile(outputFile, 2, True)

	'��s���ǂݏo��
	Dim i, tempLine, tempFields
	Do Until inputText.AtEndOfStream
'		���̓e�X�g
'		if(MsgBox(outputText.ReadLine & vbCrLf & "�I�����܂����H", vbOKCancel+vbQuestion, "CSV to VCF"))=vbOK Then
'			Wscript.Quit
'		End If
		tempLine = inputText.ReadLine
		tempFields = Split(tempLine, ",")
		tempLine = ""

		For i=0 to UBound(tempFields)
			tempLine = tempLine & tempFields(i) & vbCrLf
	        If InStr(1, tempFields(i), "END:VCARD") > 0 Then
	            Exit For
       		End If
   		Next

'		�o�̓e�X�g
'		if(MsgBox(tempLine & vbCrLf & "�I�����܂����H", vbOKCancel+vbQuestion, "CSV to VCF"))=vbOK Then
'			Wscript.Quit
'		End If
			outputText.Write tempLine
	Loop
Next

MsgBox "�I�����܂���", vbOKOnly, "CSV to VCF"

inputText.Close : outputText.Close
Set objFS=Nothing : Set objShell=Nothing
