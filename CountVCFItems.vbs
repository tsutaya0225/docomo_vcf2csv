Option Explicit

'FileSystem�I�u�W�F�N�g���쐬����
Dim objFS : Set objFS=CreateObject("Scripting.FileSystemObject")

'Shell�I�u�W�F�N�g���쐬����
Dim objShell : Set objShell=CreateObject("WScript.Shell")

'�������Ȃ���΂��̂܂܏I������
If WScript.Arguments.Count = 0 Then
	MsgBox "�t�@�C�����h���b�v���Ă��������B", vbOKOnly+vbInformation, "CountVCFItems"
	WScript.Quit
End If

Dim inputFile

'�����������̏ꍇ�͌J��Ԃ����s����
For Each inputFile In WScript.Arguments
	'�����̐�΃p�X���擾
	inputFile=LCase(objFS.getAbsolutePathName(inputFile))
	'���̓t�@�C���̊g���q���m�F
	If objFS.GetExtensionName(inputFile) <> "vcf" Then
		MsgBox "VCF�t�@�C���ł͂���܂���B" & vbCrLf & "�I�����܂��B", vbOKOnly+vbInformation, "CountVCFItems"
		Wscript.Quit
	End If

	'�e�L�X�g���擾����
	Dim inputText : set inputText= objFS.OpenTextFile(inputFile, 1)

	'��s���ǂݏo��
	Dim Counter : Counter=0
	Do Until inputText.AtEndOfStream
'		���̓e�X�g
'		if(MsgBox(outputText.ReadLine & vbCrLf & "�I�����܂����H", vbOKCancel+vbQuestion, "CountVCFItems"))=vbOK Then
'		Wscript.Quit
'		End If

		if inputText.ReadLine="BEGIN:VCARD" Then
			Counter = Counter+1
		End if
	Loop
Next

MsgBox "�f�[�^��" & Counter & "���i�[����Ă��܂��B", vbOKOnly, "CountVCFItems"

inputText.Close
Set objFS=Nothing : Set objShell=Nothing
