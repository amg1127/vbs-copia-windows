Dim fso, logf, arquivo, cod_a, cod_z, cod_letra, objdrv, letra, ppath, dty, dtlt, msgb, wshn
On Error Resume Next
Set fso = CreateObject ("Scripting.FileSystemObject")
If Err.Number = 0 Then
    Set arquivo = fso.GetFile (WScript.ScriptFullName)
    If Err.Number = 0 Then
        ppath = arquivo.ParentFolder.Path
        Set logf = fso.OpenTextFile (ppath & "\" & arquivo.Name & ".log", 8, True)
        If Err.Number = 0 Then
            msgb = ""
            Set wshn = CreateObject ("WScript.Network")
            If Err.Number <> 0 Then
                Call Err.Clear
            Else
                msgb = " [User=" & wshn.UserDomain & "\" & wshn.UserName & "; Workstation=" & wshn.ComputerName & "]"
                If Err.Number <> 0 Then
                    msgb = ""
                    Call Err.Clear
                End If
            End If
            Call RegistraLog ("Beginning this dirty job..." & msgb)
            dty = arquivo.Drive.DriveType
            If dty = 2 Then
                dtlt = UCase(arquivo.Drive.DriveLetter)
                cod_a = 65
                cod_z = 90
                For cod_letra = cod_a To cod_z
                    letra = Chr (cod_letra)
                    Call RegistraLog ("GetDrive '" & letra & ":'")
                    If letra <> dtlt Then
                        Set objdrv = fso.GetDrive (letra & ":")
                        If Not RegistraErr Then
                            If objdrv.IsReady Then
                                dty = objdrv.DriveType
                                If dty = 1 Then
                                    Call RegistraLog ("Starting copy procedure of drive '" & letra & ":'")
                                    Call CopiaItens (letra & ":", ppath & "\dest_drive_" & letra)
                                    Call RegistraLog ("Finished copy procedure of drive '" & letra & ":'")
                                Else
                                    Call RegistraLog ("WARN: Drive '" & letra & ":' is a drive with type '" & dty & ": " & drvTypeToStr(dty) & "'! Skipping...")
                                End If
                            Else
                                Call RegistraLog ("WARN: Drive '" & letra & ":' not ready! Skipping...")
                            End If
                        End If
                    Else
                        Call RegistraLog ("WARN: Source drive '" & letra & ":' is also the destination drive! Skipping...")
                    End If
                Next
                Call RegistraLog ("Finished this dirty job!")
            Else
                Call RegistraLog ("ERR: must be run under a fixed drive (running under drive type '" & dty & ": " & drvTypeToStr(dty) & "')!")
            End If
        End If
    End If
End If

Function drvTypeToStr (dty_n)
    If dty_n = 1 Then
        drvTypeToStr = "removable"
    ElseIf dty_n = 2 Then
        drvTypeToStr = "fixed"
    ElseIf dty_n = 3 Then
        drvTypeToStr = "network share"
    ElseIf dty_n = 4 Then
        drvTypeToStr = "CD-ROM"
    ElseIf dty_n = 5 Then
        drvTypeToStr = "RAM"
    Else
        drvTypeToStr = "(unknown)"
    End If
End Function

Sub RegistraLog (msg)
    Dim agora
    agora = Now
    Call logf.Write ("[" & Year(agora) & "-" & Right("0" & Month(agora), 2) & "-" & Right("0" & Day(agora), 2) & " " & _
        Right("0" & Hour(agora), 2) & ":" & Right("0" & Minute(agora), 2) & ":" & Right("0" & Second(agora), 2) & "] " & msg & vbCrLf)
End Sub

Function RegistraErr
    If Err.Number <> 0 Then
        Call RegistraLog ("ERR: " & Err.Number & " -> '" & Err.Description & "'")
        Call Err.Clear
        RegistraErr = True
    Else
        RegistraErr = False
    End If
End Function

Sub CriaPastas (orig, dest)
    Dim pas, i
    Set pas = fso.GetFolder (orig)
    For Each i In pas.SubFolders
        Call CriaPastas (i.Path, dest & "\" & i.Name)
    Next
End Sub

Sub fsoCopyFile (orig, dest)
    Dim steps, maxsteps
    On Error Resume Next
    steps = 1
    maxsteps = 4
    Do
        Call RegistraLog ("CP[" & steps & "] '" & orig & "' -> '" & dest & "'")
        Call fso.CopyFile (orig, dest, True)
        If RegistraErr Then
            steps = steps + 1
        Else
            Exit Do
        End If
    Loop While steps < maxsteps
End Sub

Sub fsoCreateFolder (dest)
    On Error Resume Next
    Call RegistraLog ("MKDIR '" & dest & "'")
    Call fso.CreateFolder (dest)
End Sub

Sub CopiaItens (orig, dest)
    Dim pas, item, colch_a, colch_f
    On Error Resume Next
    If fso.FileExists (orig) Then
        Call RegistraLog ("GetFile '" & orig & "'")
        Set pas = fso.GetFile (orig)
        If Not RegistraErr Then
            If fso.FolderExists (dest) Then
                Call RegistraLog ("WARN: '" & orig & "' is a file, but '" & dest & "' is a directory! Copying the file within the folder...")
                Call CopiaItens (orig, dest & "\" & pas.Name)
            Else
                Call fsoCopyFile (orig, dest)
            End If
        End If
    ElseIf fso.FolderExists (orig) Then
        Call RegistraLog ("GetFolder '" & orig & "'")
        Set pas = fso.GetFolder (orig)
        If Not RegistraErr Then
            If fso.FileExists (dest) Then
                Call RegistraLog ("WARN: '" & orig & "' is a directory, but '" & dest & "' is a file! Copying folder contents to another location...")
                Call CopiaItens (orig, dest & "_")
            Else
                If Not fso.FolderExists (dest) Then
                    Call fsoCreateFolder (dest)
                End If
                If fso.FolderExists (dest) Then
                    For Each item In pas.Files
                        If fso.FolderExists (orig) Then
                            Call CopiaItens (item.Path, dest & "\" & item.Name)
                        End If
                    Next
                    For Each item In pas.SubFolders
                        If fso.FolderExists (orig) Then
                            Call CopiaItens (item.Path, dest & "\" & item.Name)
                        End If
                    Next
                End If
            End If
        End If
    End If
End Sub
