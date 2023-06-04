Option Explicit

Const ForReading=1, ForWriting=2, ForAppending=8

'FileSystemオブジェクトを作成する
Dim objFS: Set objFS=CreateObject("Scripting.FileSystemObject")

'Shellオブジェクトを作成する
Dim objShell: Set objShell=CreateObject("WScript.Shell")

'引数がなければそのまま終了する
If WScript.Arguments.Count=0 Then
	MsgBox "ファイルをドロップしてください。", vbOKOnly+vbInformation, "CSV to VCF"
	WScript.Quit
End If

'main
Dim inputFile, outputFile

'引数が複数の場合は繰り返し実行する
For Each inputFile In WScript.Arguments
	'入力ファイルの拡張子を確認
	If LCase(objFS.GetExtensionName(inputFile))<>"csv" Then
		MsgBox "CSVファイルではありません。" & vbCrLf & "終了します。", vbOKOnly+vbCritical, "CSV to VCF"
		Wscript.Quit
	End If

	'出力ファイル名を設定
	outputFile=objFS.getParentFolderName(inputFile) & "\" & objFS.GetBaseName(inputFile) & ".vcf"
'	出力ファイル名の確認
'	MsgBox "outputFile は " & outputFile & " です", vbOKOnly+vbInformation, "TEST"

	'出力ファイルを作成
	If objFS.FileExists(outputFile) Then
		if(MsgBox(outputFile & " は既に存在します。" & vbCrLf & "上書きしますか？", +vbExclamation, "CSV to VCF"))=vbCancel Then
			MsgBox "終了します"
			WScript.Quit
		End If
	End If

	'テキストを取得する
	Dim inputText: set inputText= objFS.OpenTextFile(inputFile, ForReading)
	Dim outputText: Set outputText = objFS.OpenTextFile(outputFile, ForWriting, True) 'ファイルが存在しない場合は作成する

	'一行ずつ読み出す
	Dim i, tempLine, tempTexts, counter: counter=0
	Do Until inputText.AtEndOfStream
		tempLine=inputText.ReadLine
		counter=counter+1
'		入力テスト
'		if(MsgBox("入力: " & tempLine & vbCrLf & "終了しますか？", vbOKCancel+vbQuestion, "TEST"))=vbOK Then Wscript.Quit
		tempTexts=Split(tempLine, ",")
		For i=0 to UBound(tempTexts)
'			ループテスト
'			if(MsgBox("i=" & i & vbCrLf & "終了しますか？", vbOKCancel+vbQuestion, "TEST"))=vbOK Then Wscript.Quit
			iF counter=1 And i=2 Then '3行目の要素のみ "X-DCM-EXPORT:manual" を挿入
				outputText.Write "X-DCM-EXPORT:manual" & vbCrLf
				i=1 '次のループで元の要素を参照するようカウントを戻す
				counter=counter+1 '一度のみ書き込む
			Else
				outputText.Write tempTexts(i) & vbCrLf
				If InStr(1, tempTexts(i), "END:VCARD") Then Exit For
			End If
		Next
	Loop
Next

MsgBox "終了しました", vbOKOnly, "CSV to VCF"

inputText.Close: outputText.Close
Set objFS=Nothing: Set objShell=Nothing
