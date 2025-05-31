sURL = "http://www.MySite.com/test.aspx"

For i = 1 to (24*59) 'minutes in day
	
	sData = GetUrlData(sURL)
	
	If sData <> "Good" Then
		Set oShell = CreateObject("WScript.Shell")
		iButton = oShell.Popup ("Cannot connect to " & sURL, 15)
		if iButton <> -1 Then
			OpenIE sURL
		End If
		
		Log Now() & vbTab & sData
	End If

	WScript.Sleep(1000*60) '1 minute
Next

Function GetUrlData(sUrl)
	on error resume next
	Dim oHttp
	'Set oHttp = CreateObject("Microsoft.XMLHTTP")
	Set oHttp = CreateObject("MSXML2.ServerXMLHTTP")
    oHttp.setTimeouts 0, 0, 0, 0
	
	oHttp.Open "GET", sUrl, False
	oHttp.send
	
	If oHttp.Status >= 400 And oHttp.Status <= 599 And oHttp.responseText = "" Then
	  GetUrlData = "Error Occurred : " & oHttp.Status & " - " & oHttp.statusText
	Else
	  GetUrlData = oHttp.responseText
	End If
	
	Set oHttp = Nothing
End Function

Sub Log(sLine)
	Const ForAppending = 8
	Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
	Dim oLogFile: Set oLogFile = fso.OpenTextFile(WScript.ScriptFullName & ".log", ForAppending, True)
	oLogFile.WriteLine sLine
	oLogFile.Close
End Sub

Sub OpenIE(sURL)
	On Error Resume Next
	
	Set oShell = CreateObject("Shell.Application")
	Set oIE = oShell.Windows.Item

	oIE.Navigate2 sURL
	
	If Err.number <> 0 Then
		Set oIE = wscript.CreateObject("internetexplorer.application")
		oIE.Visible = True
		oIE.Navigate2 sURL
	End If
	
	Set oIE = Nothing
	Set oShell = Nothing
End Sub
