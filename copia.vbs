Dim fso, o(15), okk, ori, des
Set fso = CreateObject ("Scripting.FileSystemObject")

Sub CriaPastas (orig, dest)
	Dim pas, i
	Set pas = fso.GetFolder (orig)
	For Each i In pas.SubFolders
		Call CriaPastas (i.Path, dest & "\" & i.Name)
	Next
End Sub

Sub fsoCopyFile (orig, dest)
	Dim resp
	On Error Resume Next
	resp = 2
	Call WScript.Echo ("CP '" & orig & "' -> '" & dest & "'")
	Call fso.CopyFile (orig, dest)
	If Err.Number <> 0 Then
		If Err.Number <> 76 Then
			Call WScript.Echo ("ERR: " & Err.Number & " -> '" & Err.Description & "'")
			resp = MsgBox ("CP '" & orig & "' -> '" & dest & "'" & vbCrLf & vbCrLf & "ERR: " & Err.Number & " -> '" & Err.Description & "'", 21)
		End If
		Call Err.Clear
	End If
	If resp <> 2 Then
		Call fsoCopyFile (orig, dest)
	End If
End Sub

Sub fsoDeleteFile (dest)
	Dim resp
	On Error Resume Next
	resp = 2
	Call WScript.Echo ("DELE '" & dest & "'")
	Call fso.DeleteFile (dest)
	If Err.Number <> 0 Then
		If Err.Number <> 76 Then
			Call WScript.Echo ("ERR: " & Err.Number & " -> '" & Err.Description & "'")
			resp = MsgBox ("DELE '" & dest & "'" & vbCrLf & vbCrLf & "ERR: " & Err.Number & " -> '" & Err.Description & "'", 21)
		End If
		Call Err.Clear
	End If
	If resp <> 2 Then
		Call fsoDeleteFile (dest)
	End If
End Sub

Sub fsoDeleteFolder (dest)
	Dim resp
	On Error Resume Next
	resp = 2
	Call WScript.Echo ("RMDIR '" & dest & "'")
	Call fso.DeleteFolder (dest)
	If Err.Number <> 0 Then
		If Err.Number <> 76 Then
			Call WScript.Echo ("ERR: " & Err.Number & " -> '" & Err.Description & "'")
			resp = MsgBox ("RMDIR '" & dest & "'" & vbCrLf & vbCrLf & "ERR: " & Err.Number & " -> '" & Err.Description & "'", 21)
		End If
		Call Err.Clear
	End If
	If resp <> 2 Then
		Call fsoDeleteFolder (dest)
	End If
End Sub

Sub fsoCreateFolder (dest)
	Dim resp
	On Error Resume Next
	resp = 2
	Call WScript.Echo ("MKDIR '" & dest & "'")
	Call fso.CreateFolder (dest)
	If Err.Number <> 0 Then
		If Err.Number <> 76 Then
			Call WScript.Echo ("ERR: " & Err.Number & " -> '" & Err.Description & "'")
			resp = MsgBox ("MKDIR '" & dest & "'" & vbCrLf & vbCrLf & "ERR: " & Err.Number & " -> '" & Err.Description & "'", 21)
		End If
		Call Err.Clear
	End If
	If resp <> 2 Then
		Call fsoCreateFolder (dest)
	End If
End Sub

Sub CopiaItens (orig, dest)
	Dim pas, item
	If fso.FileExists (orig) Then
		If fso.FolderExists (dest) Then
			Call fsoDeleteFolder (dest)
		End If
		Call fsoCopyFile (orig, dest)
	ElseIf fso.FolderExists (orig) Then
		On Error Resume Next
		If fso.FileExists (dest) Then
			Call fsoDeleteFile (dest)
		End If
		If Not fso.FolderExists (dest) Then
			Call fsoCreateFolder (dest)
		End If
		Do
			On Error Resume Next
			Call WScript.Echo ("CHDIR '" & orig & "'")
			Set pas = fso.GetFolder (orig)
			If Err.Number <> 0 Then
				If Err.Number <> 76 Then
					Call WScript.Echo ("ERR: " & Err.Number & " -> '" & Err.Description & "'")
					Call MsgBox ("CHDIR '" & orig & "'" & vbCrLf & vbCrLf & "ERR: " & Err.Number & " -> '" & Err.Description & "'", 16)
				End If
				Call Err.Clear
			Else
				Exit Do
			End If
		Loop While True
		For Each item In pas.Files
			Call CopiaItens (item.Path, dest & "\" & item.Name)
		Next
		For Each item In pas.SubFolders
			Call CopiaItens (item.Path, dest & "\" & item.Name)
		Next
	End If
End Sub

o(0) = "ART4"
o(1) = "Documents and Settings"
o(2) = "DTS"
o(3) = "facil"
o(4) = "InterLattes"
o(5) = "Lattes"
o(6) = "Macromedia"
o(7) = "Meus Downloads"
o(8) = "Office52"
o(9) = "PTWIN62"
o(10) = "Recnet"
o(11) = "Roms"
o(12) = "SC2000"
o(13) = "VideoCAM Express V2"
o(14) = "Znes"

For okk = 8 To 14
	ori = "C:\" & o(okk)
	des = "D:\o que estava no C\" & o(okk)
	Call CopiaItens (ori, des)
Next
