Option Explicit

Function fnUsage()
Dim usage

usage = "" & vbCrLf & _
"Usage:" & vbCrLf & _
"    cscript Disable-SLUI.vbs -list    windir" & vbCrLf & _
"    cscript Disable-SLUI.vbs -disable windir" & vbCrLf & _
"    cscript Disable-SLUI.vbs -enable  windir" & vbCrLf & _
"" & vbCrLf & _
"*NOTE*" & vbCrLf & _
"    Do not use this in a running Windows" & vbCrLf & _
"    Boot into CMD Console using USB/DVD Windows Setup Disk" & vbCrLf & _
""
WScript.Echo usage
End Function
Function fnExit()
    WScript.Quit
End Function

Dim param_action, param_windir
If WScript.Arguments.Count = 2 Then
    param_action = WScript.Arguments(0)
    param_windir = WScript.Arguments(1)
Else
    param_action = "unknown"
End If

Select Case(LCase(param_action))
    Case "-list", "-disable", "-enable"
    Case Else
        fnUsage
        fnExit
End Select

REM USERNAME=SYSTEM
REM USERPROFILE=X:\windows\system32\config\systemprofile
Dim WshShell, WshEnv
Dim env_username, env_profile
Set WshShell = WScript.CreateObject("WScript.Shell")
Set WshEnv = WshShell.Environment("PROCESS")
env_username = WshEnv("USERNAME")
env_profile  = WshEnv("USERPROFILE")
If InStr(LCase(env_username), LCase("SYSTEM")) > 0 Or InStr(LCase(env_profile), LCase("\windows\system32\config\systemprofile")) Then
    WScript.Echo ""
    WScript.Echo "   USERNAME=" & env_username
    WScript.Echo "USERPROFILE=" & env_profile
    WScript.Echo ""
Else
    WScript.Echo ""
    WScript.Echo "***********"
    WScript.Echo "* WARNING *" & "    Seems in a running Windows, MAY NOT WORK!"
    WScript.Echo "***********"
    WScript.Echo ""
    WScript.Echo "   USERNAME=" & env_username
    WScript.Echo "USERPROFILE=" & env_profile
    WScript.Echo ""
End If

Dim fso, arrSLUI, slui
Set fso = WScript.CreateObject("Scripting.FileSystemObject")
arrSLUI = fnListAllSLUIFiles(param_windir)
For Each slui In arrSLUI
    WScript.Echo slui
    Select Case(LCase(param_action))
        Case "-list"
        Case "-disable"
            fnDisableSLUI slui
        Case "-enable"
            fnEnableSLUI slui
    End Select
Next

''
''
''
Function fnListAllSLUIFiles(ByVal full_path)
Dim arrRet(), i, n, arrTmp
Dim count, files, file, folders, folder

    On Error Resume Next
    ReDim arrRet(0)
    arrRet(0) = 0
    fnListAllSLUIFiles = arrRet

    If Not fso.FolderExists(full_path) Then
        Exit Function
    ElseIf fso.GetFolder(full_path).IsRootFolder Then
        Exit Function
    End If

    n = 0
    Set folders = fso.GetFolder(full_path).SubFolders
    count = folders.Count
    If Err.Number = 0 Then
        For Each folder In folders
            arrTmp = fnListAllSLUIFiles(folder.Path)
            If arrTmp(0) > 0 Then
                ReDim Preserve arrRet(n + arrTmp(0))
                For i = 1 To arrTmp(0)
                    arrRet(n + i ) = arrTmp(i)
                Next
                n = n + arrTmp(0)
                arrRet(0) = n
            End If
        Next
    Else
        Err.Clear
    End If

    Set files = fso.GetFolder(full_path).Files
    count = files.Count
    If Err.Number = 0 Then
        For Each file In files
            If InStr(UCase(file.Name), UCase("slui.exe")) > 0 Or InStr(UCase(file.Name), UCase("slui.exe.mui")) > 0 Then
                n = n + 1
                ReDim Preserve arrRet(n)
                arrRet(0) = n
                arrRet(n) = file.Path
            End If
        Next
    Else
        Err.Clear
    End If

    fnListAllSLUIFiles = arrRet
End Function
Function fnDisableSLUI(ByVal full_path)
Dim file

    On Error Resume Next
    Set file = fso.GetFile(full_path)
    If Err.Number = 0 Then
        If InStr(file.Name, ".disabled") > 0 Then
            WScript.Echo "    " & "already disabled"
        Else
            file.Move file.ParentFolder.Path & "\" & file.Name & ".disabled"
            If Err.Number = 0 Then
                WScript.Echo "    " & "disabled"
            Else
                WScript.Echo "    " & "disable failed"
                Err.Clear
            End If
        End If
    Else
        Err.Clear
    End If
End Function
Function fnEnableSLUI(ByVal full_path)
Dim file

    On Error Resume Next
    Set file = fso.GetFile(full_path)
    If Err.Number = 0 Then
        If InStr(file.Name, ".disabled") = 0 Then
            WScript.Echo "    " & "already enabled"
        Else
            file.Move file.ParentFolder.Path & "\" & Replace(file.Name, ".disabled", "")
            If Err.Number = 0 Then
                WScript.Echo "    " & "enabled"
            Else
                WScript.Echo "    " & "enable failed"
                Err.Clear
            End If
        End If
    Else
        Err.Clear
    End If
End Function
