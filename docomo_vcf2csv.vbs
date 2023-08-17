Option Explicit

Const ForReading=1, ForWriting=2, ForAppending=8

'FileSystemオブジェクトを作成する
Dim objFS: Set objFS=CreateObject("Scripting.FileSystemObject")

'Shellオブジェクトを作成する
Dim objShell: Set objShell=CreateObject("WScript.Shell")

'引数がなければそのまま終了する
If WScript.Arguments.Count=0 Then
	MsgBox "ファイルをドロップしてください。", vbOKOnly+vbInformation, "VCF to CSV"
	WScript.Quit
End If

'main
Dim inputFile, outputFile

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
			WScript.Quit
		End If
	End If

	'テキストを取得する
	Dim inputText: set inputText=objFS.OpenTextFile(inputFile, ForReading)
	Dim outputText: Set outputText=objFS.OpenTextFile(outputFile, ForWriting, True) 'ファイルが存在しない場合は作成する

	'一行ずつ読み出す
	Dim tempField, counter: counter=0
	Do Until inputText.AtEndOfStream
		Do
			tempField=inputText.ReadLine
			counter=counter+1
			'3行目の要素のみ "X-DCM-EXPORT:manual" なので無視
			if counter=3 Then
				'何もしない
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

MsgBox "終了しました", vbOKOnly, "VCF to CSV"

inputText.Close: outputText.Close
Set objFS=Nothing: Set objShell=Nothing
