Option Explicit

'FileSystem�I�u�W�F�N�g���쐬����
Dim objFS : Set objFS=CreateObject("Scripting.FileSystemObject")

'Shell�I�u�W�F�N�g���쐬����
Dim objShell : Set objShell=CreateObject("WScript.Shell")

'�������Ȃ���΂��̂܂܏I������
If WScript.Arguments.Count = 0 Then
	MsgBox "�t�@�C�����h���b�v���Ă��������B", vbOKOnly+vbInformation, "VCF to CSV"
	WScript.Quit
End If

'main
Dim inputFile, outputFile

'�����������̏ꍇ�͌J��Ԃ����s����
For Each inputFile In WScript.Arguments
	'�����̐�΃p�X���擾
	inputFile=LCase(objFS.getAbsolutePathName(inputFile))
	'���̓t�@�C���̊g���q���m�F
	If objFS.GetExtensionName(inputFile) <> "vcf" Then
	MsgBox "VCF�t�@�C���ł͂���܂���B" & vbCrLf & "�I�����܂��B", vbOKOnly+vbInformation, "VCF to CSV"
	Wscript.Quit
	End If

	'�o�̓t�@�C������ݒ�
	outputFile=objFS.getParentFolderName(inputFile) & "\" & objFS.GetBaseName(inputFile) & ".csv"
'	�o�̓t�@�C�����̊m�F
'	MsgBox "outputFile�� " & outputFile & " �ł�", vbOKOnly+vbInformation, "VCF to CSV"
	'�o�̓t�@�C�����쐬
	If objFS.FileExists(outputFile) Then
		if(MsgBox(outputFile & "�͊��ɑ��݂��܂��B" & vbCrLf & "�㏑�����܂����H", +vbExclamation, "VCF to CSV"))=vbCancel Then
			MsgBox "�I�����܂�"
			WScript.Quit
		End If
	End If

	'�e�L�X�g���擾����
	Dim inputText : set inputText= objFS.OpenTextFile(inputFile, 1)
	Dim outputText : Set outputText = objFS.OpenTextFile(outputFile, 2, True)

	'��s���ǂݏo��
	Do Until inputText.AtEndOfStream
'		���̓e�X�g
'		if(MsgBox(outputText.ReadLine & vbCrLf & "�I�����܂����H", vbOKCancel+vbQuestion, "VCF to CSV"))=vbOK Then
'			Wscript.Quit
'		End If
		Dim tempText
		Do
			tempText = inputText.ReadLine
			if tempText = "X-DCM-EXPORT:manual" Then Exit Do
			outputText.Write tempText
			if tempText="END:VCARD" Then
				outputText.Write vbCrLf
				Exit Do
			Else
				outputText.Write ","
			End if
		Loop
	Loop
Next

MsgBox "�I�����܂���", vbOKOnly, "VCF to CSV"

inputText.Close : outputText.Close
Set objFS=Nothing : Set objShell=Nothing
