Option Explicit

'FileSystemオブジェクトを作成する
Dim objFS : Set objFS=CreateObject("Scripting.FileSystemObject")

'Shellオブジェクトを作成する
Dim objShell : Set objShell=CreateObject("WScript.Shell")

'引数がなければそのまま終了する
If WScript.Arguments.Count = 0 Then
	MsgBox "ファイルをドロップしてください。", vbOKOnly+vbInformation, "CountVCFItems"
	WScript.Quit
End If

Dim inputFile

'引数が複数の場合は繰り返し実行する
For Each inputFile In WScript.Arguments
	'引数の絶対パスを取得
	inputFile=LCase(objFS.getAbsolutePathName(inputFile))
	'入力ファイルの拡張子を確認
	If objFS.GetExtensionName(inputFile) <> "vcf" Then
		MsgBox "VCFファイルではありません。" & vbCrLf & "終了します。", vbOKOnly+vbInformation, "CountVCFItems"
		Wscript.Quit
	End If

	'テキストを取得する
	Dim inputText : set inputText= objFS.OpenTextFile(inputFile, 1)

	'一行ずつ読み出す
	Dim Counter : Counter=0
	Do Until inputText.AtEndOfStream
'		入力テスト
'		if(MsgBox(outputText.ReadLine & vbCrLf & "終了しますか？", vbOKCancel+vbQuestion, "CountVCFItems"))=vbOK Then
'		Wscript.Quit
'		End If

		if inputText.ReadLine="BEGIN:VCARD" Then
			Counter = Counter+1
		End if
	Loop
Next

MsgBox "データは" & Counter & "件格納されています。", vbOKOnly, "CountVCFItems"

inputText.Close
Set objFS=Nothing : Set objShell=Nothing
