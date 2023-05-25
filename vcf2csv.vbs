Option Explicit

'FileSystemオブジェクトを作成する
Dim objFS : Set objFS=CreateObject("Scripting.FileSystemObject")

'Shellオブジェクトを作成する
Dim objShell : Set objShell=CreateObject("WScript.Shell")

'引数がなければそのまま終了する
If WScript.Arguments.Count = 0 Then
	MsgBox "ファイルをドロップしてください。", vbOKOnly+vbInformation, "VCF to CSV"
	WScript.Quit
End If

'main
Dim inputFile, outputFile

'引数が複数の場合は繰り返し実行する
For Each inputFile In WScript.Arguments
	'引数の絶対パスを取得
	inputFile=LCase(objFS.getAbsolutePathName(inputFile))
	'入力ファイルの拡張子を確認
	If objFS.GetExtensionName(inputFile) <> "vcf" Then
	MsgBox "VCFファイルではありません。" & vbCrLf & "終了します。", vbOKOnly+vbInformation, "VCF to CSV"
	Wscript.Quit
	End If

	'出力ファイル名を設定
	outputFile=objFS.getParentFolderName(inputFile) & "\" & objFS.GetBaseName(inputFile) & ".csv"
'	出力ファイル名の確認
'	MsgBox "outputFileは " & outputFile & " です", vbOKOnly+vbInformation, "VCF to CSV"
	'出力ファイルを作成
	If objFS.FileExists(outputFile) Then
		if(MsgBox(outputFile & "は既に存在します。" & vbCrLf & "上書きしますか？", +vbExclamation, "VCF to CSV"))=vbCancel Then
			MsgBox "終了します"
			WScript.Quit
		End If
	End If

	'テキストを取得する
	Dim inputText : set inputText= objFS.OpenTextFile(inputFile, 1)
	Dim outputText : Set outputText = objFS.OpenTextFile(outputFile, 2, True)

	'一行ずつ読み出す
	Do Until inputText.AtEndOfStream
'		入力テスト
'		if(MsgBox(outputText.ReadLine & vbCrLf & "終了しますか？", vbOKCancel+vbQuestion, "VCF to CSV"))=vbOK Then
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

MsgBox "終了しました", vbOKOnly, "VCF to CSV"

inputText.Close : outputText.Close
Set objFS=Nothing : Set objShell=Nothing
