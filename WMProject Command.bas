Attribute VB_Name = "Module1"
'By WMProject1217
Sub Main()
'dim and set vals
On Error GoTo Error
Dim a, b, c, d, e, f, g, h, i, j, k, l, m, n, o, p, q, r, s, t, u, v, w, x, y, z
Dim tmpa As String, tmpb() As String, tmpc As String, tmpd() As String, tempe, tempf, tempg
Dim cmd, sysdisk, sysroot, retval, path
cmd = Command
sysdisk = Environ("SystemDrive")
sysroot = Environ("SystemRoot")
'check cmd
If cmd = "" Then
Exit Sub
End If
tempe = InStr(1, cmd, "=")
If tempe = 0 Then
Exit Sub
End If
'split cmd
cmd = Replace(cmd, "#AppPath#", App.path)
cmd = Replace(cmd, "#SystemRoot#", sysroot)
cmd = Replace(cmd, "#SystemDrive#", sysdisk)
cmd = Replace(cmd, "#Date#", Date)
cmd = Replace(cmd, "#Time#", Time)
tmpb() = Split(cmd, "=")
If tmpb(0) = "WMDEBUG" Then

End If
'Command EXEC
If tmpb(0) = "exec" Or tmpb(0) = "EXEC" Or tmpb(0) = "Exec" Then
path = tmpb(1)
retval = Shell(path, vbNormalFocus)
Exit Sub
End If
'Command EXECHC
If tmpb(0) = "exechc" Or tmpb(0) = "EXECHC" Or tmpb(0) = "Exechc" Then
path = tmpb(1)
retval = Shell(path, vbHide)
Exit Sub
End If
'Command KILL
If tmpb(0) = "kill" Or tmpb(0) = "KILL" Or tmpb(0) = "Kill" Then
path = sysroot & "\System32\taskkill.exe /f /im " & tmpb(1)
retval = Shell(path, vbHide)
Exit Sub
End If
'Command PLAYSOUND
If tmpb(0) = "playsound" Or tmpb(0) = "PLAYSOUND" Or tmpb(0) = "Playsound" Then
path = App.path & "\BGM.exe " & tmpb(1)
retval = Shell(path, vbHide)
Exit Sub
End If
'Command MSGBOX
If tmpb(0) = "msgbox" Or tmpb(0) = "MSGBOX" Or tmpb(0) = "Msgbox" Then
path = App.path & "\MSGBOX.exe " & tmpb(1) & "=" & tmpb(2)
retval = Shell(path, vbNormalFocus)
Exit Sub
End If
'Command MSGLINE
If tmpb(0) = "msgline" Or tmpb(0) = "MSGLINE" Or tmpb(0) = "Msgline" Then
path = App.path & "\MSGLINE.exe " & tmpb(1) & "=" & tmpb(2) & "=" & tmpb(3) & "=" & tmpb(4) & "=" & tmpb(5) & "=" & tmpb(6) & "=" & tmpb(7) & "=" & tmpb(8) & "=" & tmpb(9) & "=" & tmpb(10) & "=" & tmpb(11) & "=" & tmpb(12) & "=" & tmpb(13) & "=" & tmpb(14) & "=" & tmpb(15) & "=" & tmpb(16) & "=" & tmpb(17)
retval = Shell(path, vbNormalFocus)
Exit Sub
End If
'Command BSDC
If tmpb(0) = "bsdc" Or tmpb(0) = "BSDC" Or tmpb(0) = "Bsdc" Then
path = App.path & "\BSDC.exe " & tmpb(1) & "=" & tmpb(2) & "=" & tmpb(3) & "=" & tmpb(4) & "=" & tmpb(5)
retval = Shell(path, vbNormalFocus)
Exit Sub
End If
'Command BSDD
If tmpb(0) = "bsdd" Or tmpb(0) = "BSDD" Or tmpb(0) = "Bsdd" Then
path = App.path & "\BSDD.exe " & tmpb(1) & "=" & tmpb(2) & "=" & tmpb(3) & "=" & tmpb(4) & "=" & tmpb(5)
retval = Shell(path, vbNormalFocus)
Exit Sub
End If
Error:
Exit Sub
End Sub
