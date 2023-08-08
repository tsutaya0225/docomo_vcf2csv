Option Explicit

Const ForReading=1, ForWriting=2, ForAppending=8

'FileSystem�I�u�W�F�N�g���쐬����
Dim objFS: Set objFS=CreateObject("Scripting.FileSystemObject")

'Shell�I�u�W�F�N�g���쐬����
Dim objShell: Set objShell=CreateObject("WScript.Shell")

'�������Ȃ���΂��̂܂܏I������
If WScript.Arguments.Count=0 Then
	MsgBox "�t�@�C�����h���b�v���Ă��������B", vbOKOnly+vbInformation, "VCF to CSV"
	WScript.Quit
End If

'main
Dim inputFile, outputFile

'�����������̏ꍇ�͌J��Ԃ����s����
For Each inputFile In WScript.Arguments
	'���̓t�@�C���̊g���q���m�F
	If LCase(objFS.GetExtensionName(inputFile))<>"vcf" Then
		MsgBox "VCF�t�@�C���ł͂���܂���B" & vbCrLf & "�I�����܂��B", vbOKOnly+vbInformation, "VCF to CSV"
		WScript.Quit
	End If

	'�o�̓t�@�C������ݒ�
	outputFile=objFS.getParentFolderName(inputFile) & "\" & objFS.GetBaseName(inputFile) & ".csv"
'	�o�̓t�@�C�����̊m�F
'	MsgBox "outputFile �� " & outputFile & " �ł�", vbOKOnly+vbInformation, "TEST"

	'�o�̓t�@�C�����쐬
	If objFS.FileExists(outputFile) Then
		if(MsgBox(outputFile & " �͊��ɑ��݂��܂��B" & vbCrLf & "�㏑�����܂����H", +vbExclamation, "VCF to CSV"))=vbCancel Then
			MsgBox "�I�����܂�", vbOKOnly+vbInformation, "VCF to CSV"
			WScript.Quit
		End If
	End If

	'�e�L�X�g���擾����
	Dim inputText: set inputText=objFS.OpenTextFile(inputFile, ForReading)
	Dim outputText: Set outputText=objFS.OpenTextFile(outputFile, ForWriting, True) '�t�@�C�������݂��Ȃ��ꍇ�͍쐬����

	'��s���ǂݏo��
	Dim tempField, counter: counter=0
	Do Until inputText.AtEndOfStream
		Do
			tempField=inputText.ReadLine
			counter=counter+1
			'3�s�ڂ̗v�f�̂� "X-DCM-EXPORT:manual" �Ȃ̂Ŗ���
			if counter=3 Then
				'�������Ȃ�
			Else
				outputText.Write tempField
				if tempField="END:VCARD" Then
					outputText.Write vbCrLf
					Exit Do
				Else
					outputText.Write ","
				End if
			End if
		Loop
	Loop
Next

MsgBox "�I�����܂���", vbOKOnly, "VCF to CSV"

inputText.Close: outputText.Close
Set objFS=Nothing: Set objShell=Nothing
