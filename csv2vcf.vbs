Option Explicit

Const ForReading=1, ForWriting=2, ForAppending=8

'FileSystem�I�u�W�F�N�g���쐬����
Dim objFS: Set objFS=CreateObject("Scripting.FileSystemObject")

'Shell�I�u�W�F�N�g���쐬����
Dim objShell: Set objShell=CreateObject("WScript.Shell")

'�������Ȃ���΂��̂܂܏I������
If WScript.Arguments.Count=0 Then
	MsgBox "�t�@�C�����h���b�v���Ă��������B", vbOKOnly+vbInformation, "CSV to VCF"
	WScript.Quit
End If

'main
Dim inputFile, outputFile

'�����������̏ꍇ�͌J��Ԃ����s����
For Each inputFile In WScript.Arguments
	'���̓t�@�C���̊g���q���m�F
	If LCase(objFS.GetExtensionName(inputFile))<>"csv" Then
		MsgBox "CSV�t�@�C���ł͂���܂���B" & vbCrLf & "�I�����܂��B", vbOKOnly+vbCritical, "CSV to VCF"
		Wscript.Quit
	End If

	'�o�̓t�@�C������ݒ�
	outputFile=objFS.getParentFolderName(inputFile) & "\" & objFS.GetBaseName(inputFile) & ".vcf"
'	�o�̓t�@�C�����̊m�F
'	MsgBox "outputFile �� " & outputFile & " �ł�", vbOKOnly+vbInformation, "TEST"

	'�o�̓t�@�C�����쐬
	If objFS.FileExists(outputFile) Then
		if(MsgBox(outputFile & " �͊��ɑ��݂��܂��B" & vbCrLf & "�㏑�����܂����H", +vbExclamation, "CSV to VCF"))=vbCancel Then
			MsgBox "�I�����܂�"
			WScript.Quit
		End If
	End If

	'�e�L�X�g���擾����
	Dim inputText: set inputText= objFS.OpenTextFile(inputFile, ForReading)
	Dim outputText: Set outputText = objFS.OpenTextFile(outputFile, ForWriting, True) '�t�@�C�������݂��Ȃ��ꍇ�͍쐬����

	'��s���ǂݏo��
	Dim i, tempLine, tempTexts, counter: counter=0
	Do Until inputText.AtEndOfStream
		tempLine=inputText.ReadLine
		counter=counter+1
'		���̓e�X�g
'		if(MsgBox("����: " & tempLine & vbCrLf & "�I�����܂����H", vbOKCancel+vbQuestion, "TEST"))=vbOK Then Wscript.Quit
		tempTexts=Split(tempLine, ",")
		For i=0 to UBound(tempTexts)
'			���[�v�e�X�g
'			if(MsgBox("i=" & i & vbCrLf & "�I�����܂����H", vbOKCancel+vbQuestion, "TEST"))=vbOK Then Wscript.Quit
			iF counter=1 And i=2 Then '3�s�ڂ̗v�f�̂� "X-DCM-EXPORT:manual" ��}��
				outputText.Write "X-DCM-EXPORT:manual" & vbCrLf
				i=1 '���̃��[�v�Ō��̗v�f���Q�Ƃ���悤�J�E���g��߂�
				counter=counter+1 '��x�̂ݏ�������
			Else
				outputText.Write tempTexts(i) & vbCrLf
				If InStr(1, tempTexts(i), "END:VCARD") Then Exit For
			End If
		Next
	Loop
Next

MsgBox "�I�����܂���", vbOKOnly, "CSV to VCF"

inputText.Close: outputText.Close
Set objFS=Nothing: Set objShell=Nothing
