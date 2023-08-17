Option Explicit

Const ForReading=1, ForWriting=2, ForAppending=8

<<<<<<< HEAD
'FileSystemオブジェクトを作成する
Dim objFS: Set objFS=CreateObject("Scripting.FileSystemObject")

'Shellオブジェクトを作成する
Dim objShell: Set objShell=CreateObject("WScript.Shell")

'引数がなければそのまま終了する
If WScript.Arguments.Count=0 Then
	MsgBox "ファイルをドロップしてください。", vbOKOnly+vbInformation, "VCF to CSV"
=======
'FileSystem繧ｪ繝悶ず繧ｧ繧ｯ繝医ｒ菴懈�舌☆繧�
Dim objFS: Set objFS=CreateObject("Scripting.FileSystemObject")

'Shell繧ｪ繝悶ず繧ｧ繧ｯ繝医ｒ菴懈�舌☆繧�
Dim objShell: Set objShell=CreateObject("WScript.Shell")

'蠑墓焚縺後↑縺代ｌ縺ｰ縺昴�ｮ縺ｾ縺ｾ邨ゆｺ�縺吶ｋ
If WScript.Arguments.Count=0 Then
	MsgBox "繝輔ぃ繧､繝ｫ繧偵ラ繝ｭ繝�繝励＠縺ｦ縺上□縺輔＞縲�", vbOKOnly+vbInformation, "VCF to CSV"
>>>>>>> 8f573b4bcf66654d464704d5ccff0fe26faf9579
	WScript.Quit
End If

'main
Dim inputFile, outputFile

<<<<<<< HEAD
'引数が複数の場合は繰り返し実行する
For Each inputFile In WScript.Arguments
	'入力ファイルの拡張子を確認
	If LCase(objFS.GetExtensionName(inputFile))<>"vcf" Then
		MsgBox "VCFファイルではありません。" & vbCrLf & "終了します。", vbOKOnly+vbInformation, "VCF to CSV"
		WScript.Quit
	End If

	'出力ファイル名を設定
	outputFile=objFS.getParentFolderName(inputFile) & "\" & objFS.GetBaseName(inputFile) & ".csv"
'	出力ファイル名の確認
'	MsgBox "outputFile は " & outputFile & " です", vbOKOnly+vbInformation, "TEST"

	'出力ファイルを作成
	If objFS.FileExists(outputFile) Then
		if(MsgBox(outputFile & " は既に存在します。" & vbCrLf & "上書きしますか？", +vbExclamation, "VCF to CSV"))=vbCancel Then
			MsgBox "終了します", vbOKOnly+vbInformation, "VCF to CSV"
=======
'蠑墓焚縺瑚､�謨ｰ縺ｮ蝣ｴ蜷医�ｯ郢ｰ繧願ｿ斐＠螳溯｡後☆繧�
For Each inputFile In WScript.Arguments
	'蜈･蜉帙ヵ繧｡繧､繝ｫ縺ｮ諡｡蠑ｵ蟄舌ｒ遒ｺ隱�
	If LCase(objFS.GetExtensionName(inputFile))<>"vcf" Then
		MsgBox "VCF繝輔ぃ繧､繝ｫ縺ｧ縺ｯ縺ゅｊ縺ｾ縺帙ｓ縲�" & vbCrLf & "邨ゆｺ�縺励∪縺吶��", vbOKOnly+vbInformation, "VCF to CSV"
		WScript.Quit
	End If

	'蜃ｺ蜉帙ヵ繧｡繧､繝ｫ蜷阪ｒ險ｭ螳�
	outputFile=objFS.getParentFolderName(inputFile) & "\" & objFS.GetBaseName(inputFile) & ".csv"
'	蜃ｺ蜉帙ヵ繧｡繧､繝ｫ蜷阪�ｮ遒ｺ隱�
'	MsgBox "outputFile 縺ｯ " & outputFile & " 縺ｧ縺�", vbOKOnly+vbInformation, "TEST"

	'蜃ｺ蜉帙ヵ繧｡繧､繝ｫ繧剃ｽ懈��
	If objFS.FileExists(outputFile) Then
		if(MsgBox(outputFile & " 縺ｯ譌｢縺ｫ蟄伜惠縺励∪縺吶��" & vbCrLf & "荳頑嶌縺阪＠縺ｾ縺吶°�ｼ�", +vbExclamation, "VCF to CSV"))=vbCancel Then
			MsgBox "邨ゆｺ�縺励∪縺�", vbOKOnly+vbInformation, "VCF to CSV"
>>>>>>> 8f573b4bcf66654d464704d5ccff0fe26faf9579
			WScript.Quit
		End If
	End If

<<<<<<< HEAD
	'テキストを取得する
	Dim inputText: set inputText=objFS.OpenTextFile(inputFile, ForReading)
	Dim outputText: Set outputText=objFS.OpenTextFile(outputFile, ForWriting, True) 'ファイルが存在しない場合は作成する

	'一行ずつ読み出す
=======
	'繝�繧ｭ繧ｹ繝医ｒ蜿門ｾ励☆繧�
	Dim inputText: set inputText=objFS.OpenTextFile(inputFile, ForReading)
	Dim outputText: Set outputText=objFS.OpenTextFile(outputFile, ForWriting, True) '繝輔ぃ繧､繝ｫ縺悟ｭ伜惠縺励↑縺�蝣ｴ蜷医�ｯ菴懈�舌☆繧�

	'荳�陦後★縺､隱ｭ縺ｿ蜃ｺ縺�
>>>>>>> 8f573b4bcf66654d464704d5ccff0fe26faf9579
	Dim tempField, counter: counter=0
	Do Until inputText.AtEndOfStream
		Do
			tempField=inputText.ReadLine
			counter=counter+1
<<<<<<< HEAD
			'3行目の要素のみ "X-DCM-EXPORT:manual" なので無視
			if counter=3 Then
				'何もしない
=======
			'3陦檎岼縺ｮ隕∫ｴ�縺ｮ縺ｿ "X-DCM-EXPORT:manual" 縺ｪ縺ｮ縺ｧ辟｡隕�
			if counter=3 Then
				'菴輔ｂ縺励↑縺�
>>>>>>> 8f573b4bcf66654d464704d5ccff0fe26faf9579
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

<<<<<<< HEAD
MsgBox "終了しました", vbOKOnly, "VCF to CSV"
=======
MsgBox "邨ゆｺ�縺励∪縺励◆", vbOKOnly, "VCF to CSV"
>>>>>>> 8f573b4bcf66654d464704d5ccff0fe26faf9579

inputText.Close: outputText.Close
Set objFS=Nothing: Set objShell=Nothing
