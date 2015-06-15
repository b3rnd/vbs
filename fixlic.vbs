'fix Office 365 reactivation issues: 
' https://technet.microsoft.com/en-us/library/gg702620.aspx
' http://community.spiceworks.com/how_to/48973-remove-and-re-add-license-key-for-office-2013-on-office-365

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = WScript.CreateObject("WScript.Shell")

If objFSO.FileExists("%ProgramFiles%\Microsoft Office\Office15\ospp.vbs") Then
	ospp = "%ProgramFiles%\Microsoft Office\Office15\ospp.vbs"
ElseIf  objFSO.FileExists("%ProgramFiles(x86)%\Microsoft Office\Office15\ospp.vbs") Then
	ospp = "%ProgramFiles(x86)%\Microsoft Office\Office15\ospp.vbs"
Else
	Messagebox "Office is not installed!"
	WScript.Quit
End If

Set objExec = objShell.Exec("cscript """ + ospp + """ /dstatus")

Do
	line = objExec.StdOut.ReadLine()
	If line = "LICENSE NAME: Office 15, OfficeO365ProPlusR_Subscription1 edition" Or line = "LICENSE NAME: Office 15, OfficeO365ProPlusR_Grace edition" Then
		key = ""
		Do
			line = objExec.StdOut.ReadLine()
			If left(line, 17) = "Last 5 characters" then
				key = Right(line, 5)
			End If
		Loop While key = ""
		
		Set objExec2 = objShell.Exec("cscript """ + ospp + """ /unpkey:" + key)
		
		do
			line2 = objExec2.StdOut.ReadLine()
			s = s + line2 + vbcrlf
		loop while not objExec2.Stdout.atEndOfStream
		MsgBox s
	End If
Loop While Not objExec.Stdout.atEndOfStream
