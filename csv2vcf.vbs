Option Explicit

'FileSystemオブジェクトを作成する
Dim objFS : Set objFS=CreateObject("Scripting.FileSystemObject")

'Shellオブジェクトを作成する
Dim objShell : Set objShell=CreateObject("WScript.Shell")

'引数がなければそのまま終了する
If WScript.Arguments.Count = 0 Then
	MsgBox "ファイルをドロップしてください。", vbOKOnly+vbInformation, "CSV to VCF"
	WScript.Quit
End If

Dim inputFile, outputFile

'引数が複数の場合は繰り返し実行する
For Each inputFile In WScript.Arguments
	'引数の絶対パスを取得
	inputFile=LCase(objFS.getAbsolutePathName(inputFile))
	'入力ファイルの拡張子を確認
	If objFS.GetExtensionName(inputFile) <> "csv" Then
		MsgBox "CSVファイルではありません。" & vbCrLf & "終了します。", vbOKOnly+vbCritical, "CSV to VCF"
		Wscript.Quit
	End If

	'出力ファイル名を設定
	outputFile=objFS.getParentFolderName(inputFile) & "\" & objFS.GetBaseName(inputFile) & ".vcf"
'	出力ファイル名の確認
'	MsgBox "outputFileは " & outputFile & " です", vbOKOnly+vbInformation, "CSV to VCF"
	'出力ファイルを作成
	If objFS.FileExists(outputFile) Then
		if(MsgBox(outputFile & "は既に存在します。" & vbCrLf & "上書きしますか？", +vbExclamation, "CSV to VCF"))=vbCancel Then
			MsgBox "終了します"
			WScript.Quit
		End If
	End If

	'テキストを取得する
	Dim inputText : set inputText= objFS.OpenTextFile(inputFile, 1)
	Dim outputText : Set outputText = objFS.OpenTextFile(outputFile, 2, True)

	'一行ずつ読み出す
	Dim i, tempLine, tempFields
	Do Until inputText.AtEndOfStream
'		入力テスト
'		if(MsgBox(outputText.ReadLine & vbCrLf & "終了しますか？", vbOKCancel+vbQuestion, "CSV to VCF"))=vbOK Then
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

'		出力テスト
'		if(MsgBox(tempLine & vbCrLf & "終了しますか？", vbOKCancel+vbQuestion, "CSV to VCF"))=vbOK Then
'			Wscript.Quit
'		End If
			outputText.Write tempLine
	Loop
Next

MsgBox "終了しました", vbOKOnly, "CSV to VCF"

inputText.Close : outputText.Close
Set objFS=Nothing : Set objShell=Nothing
