Option Explicit

Const ForReading=1, ForWriting=2, ForAppending=8

<<<<<<< HEAD
'FileSystemƒIƒuƒWƒFƒNƒg‚ğì¬‚·‚é
Dim objFS: Set objFS=CreateObject("Scripting.FileSystemObject")

'ShellƒIƒuƒWƒFƒNƒg‚ğì¬‚·‚é
Dim objShell: Set objShell=CreateObject("WScript.Shell")

'ˆø”‚ª‚È‚¯‚ê‚Î‚»‚Ì‚Ü‚ÜI—¹‚·‚é
If WScript.Arguments.Count=0 Then
	MsgBox "ƒtƒ@ƒCƒ‹‚ğƒhƒƒbƒv‚µ‚Ä‚­‚¾‚³‚¢B", vbOKOnly+vbInformation, "VCF to CSV"
=======
'FileSystemã‚ªãƒ–ã‚¸ã‚§ã‚¯ãƒˆã‚’ä½œæˆã™ã‚‹
Dim objFS: Set objFS=CreateObject("Scripting.FileSystemObject")

'Shellã‚ªãƒ–ã‚¸ã‚§ã‚¯ãƒˆã‚’ä½œæˆã™ã‚‹
Dim objShell: Set objShell=CreateObject("WScript.Shell")

'å¼•æ•°ãŒãªã‘ã‚Œã°ãã®ã¾ã¾çµ‚äº†ã™ã‚‹
If WScript.Arguments.Count=0 Then
	MsgBox "ãƒ•ã‚¡ã‚¤ãƒ«ã‚’ãƒ‰ãƒ­ãƒƒãƒ—ã—ã¦ãã ã•ã„ã€‚", vbOKOnly+vbInformation, "VCF to CSV"
>>>>>>> 8f573b4bcf66654d464704d5ccff0fe26faf9579
	WScript.Quit
End If

'main
Dim inputFile, outputFile

<<<<<<< HEAD
'ˆø”‚ª•¡”‚Ìê‡‚ÍŒJ‚è•Ô‚µÀs‚·‚é
For Each inputFile In WScript.Arguments
	'“ü—Íƒtƒ@ƒCƒ‹‚ÌŠg’£q‚ğŠm”F
	If LCase(objFS.GetExtensionName(inputFile))<>"vcf" Then
		MsgBox "VCFƒtƒ@ƒCƒ‹‚Å‚Í‚ ‚è‚Ü‚¹‚ñB" & vbCrLf & "I—¹‚µ‚Ü‚·B", vbOKOnly+vbInformation, "VCF to CSV"
		WScript.Quit
	End If

	'o—Íƒtƒ@ƒCƒ‹–¼‚ğİ’è
	outputFile=objFS.getParentFolderName(inputFile) & "\" & objFS.GetBaseName(inputFile) & ".csv"
'	o—Íƒtƒ@ƒCƒ‹–¼‚ÌŠm”F
'	MsgBox "outputFile ‚Í " & outputFile & " ‚Å‚·", vbOKOnly+vbInformation, "TEST"

	'o—Íƒtƒ@ƒCƒ‹‚ğì¬
	If objFS.FileExists(outputFile) Then
		if(MsgBox(outputFile & " ‚ÍŠù‚É‘¶İ‚µ‚Ü‚·B" & vbCrLf & "ã‘‚«‚µ‚Ü‚·‚©H", +vbExclamation, "VCF to CSV"))=vbCancel Then
			MsgBox "I—¹‚µ‚Ü‚·", vbOKOnly+vbInformation, "VCF to CSV"
=======
'å¼•æ•°ãŒè¤‡æ•°ã®å ´åˆã¯ç¹°ã‚Šè¿”ã—å®Ÿè¡Œã™ã‚‹
For Each inputFile In WScript.Arguments
	'å…¥åŠ›ãƒ•ã‚¡ã‚¤ãƒ«ã®æ‹¡å¼µå­ã‚’ç¢ºèª
	If LCase(objFS.GetExtensionName(inputFile))<>"vcf" Then
		MsgBox "VCFãƒ•ã‚¡ã‚¤ãƒ«ã§ã¯ã‚ã‚Šã¾ã›ã‚“ã€‚" & vbCrLf & "çµ‚äº†ã—ã¾ã™ã€‚", vbOKOnly+vbInformation, "VCF to CSV"
		WScript.Quit
	End If

	'å‡ºåŠ›ãƒ•ã‚¡ã‚¤ãƒ«åã‚’è¨­å®š
	outputFile=objFS.getParentFolderName(inputFile) & "\" & objFS.GetBaseName(inputFile) & ".csv"
'	å‡ºåŠ›ãƒ•ã‚¡ã‚¤ãƒ«åã®ç¢ºèª
'	MsgBox "outputFile ã¯ " & outputFile & " ã§ã™", vbOKOnly+vbInformation, "TEST"

	'å‡ºåŠ›ãƒ•ã‚¡ã‚¤ãƒ«ã‚’ä½œæˆ
	If objFS.FileExists(outputFile) Then
		if(MsgBox(outputFile & " ã¯æ—¢ã«å­˜åœ¨ã—ã¾ã™ã€‚" & vbCrLf & "ä¸Šæ›¸ãã—ã¾ã™ã‹ï¼Ÿ", +vbExclamation, "VCF to CSV"))=vbCancel Then
			MsgBox "çµ‚äº†ã—ã¾ã™", vbOKOnly+vbInformation, "VCF to CSV"
>>>>>>> 8f573b4bcf66654d464704d5ccff0fe26faf9579
			WScript.Quit
		End If
	End If

<<<<<<< HEAD
	'ƒeƒLƒXƒg‚ğæ“¾‚·‚é
	Dim inputText: set inputText=objFS.OpenTextFile(inputFile, ForReading)
	Dim outputText: Set outputText=objFS.OpenTextFile(outputFile, ForWriting, True) 'ƒtƒ@ƒCƒ‹‚ª‘¶İ‚µ‚È‚¢ê‡‚Íì¬‚·‚é

	'ˆês‚¸‚Â“Ç‚İo‚·
=======
	'ãƒ†ã‚­ã‚¹ãƒˆã‚’å–å¾—ã™ã‚‹
	Dim inputText: set inputText=objFS.OpenTextFile(inputFile, ForReading)
	Dim outputText: Set outputText=objFS.OpenTextFile(outputFile, ForWriting, True) 'ãƒ•ã‚¡ã‚¤ãƒ«ãŒå­˜åœ¨ã—ãªã„å ´åˆã¯ä½œæˆã™ã‚‹

	'ä¸€è¡Œãšã¤èª­ã¿å‡ºã™
>>>>>>> 8f573b4bcf66654d464704d5ccff0fe26faf9579
	Dim tempField, counter: counter=0
	Do Until inputText.AtEndOfStream
		Do
			tempField=inputText.ReadLine
			counter=counter+1
<<<<<<< HEAD
			'3s–Ú‚Ì—v‘f‚Ì‚İ "X-DCM-EXPORT:manual" ‚È‚Ì‚Å–³‹
			if counter=3 Then
				'‰½‚à‚µ‚È‚¢
=======
			'3è¡Œç›®ã®è¦ç´ ã®ã¿ "X-DCM-EXPORT:manual" ãªã®ã§ç„¡è¦–
			if counter=3 Then
				'ä½•ã‚‚ã—ãªã„
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
MsgBox "I—¹‚µ‚Ü‚µ‚½", vbOKOnly, "VCF to CSV"
=======
MsgBox "çµ‚äº†ã—ã¾ã—ãŸ", vbOKOnly, "VCF to CSV"
>>>>>>> 8f573b4bcf66654d464704d5ccff0fe26faf9579

inputText.Close: outputText.Close
Set objFS=Nothing: Set objShell=Nothing
