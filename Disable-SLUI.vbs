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
"    Use this in a running Windows is not encouraged, does not always work" & vbCrLf & _
"    Boot into CMD Console using USB/DVD Windows Setup Disk is preferred" & vbCrLf & _
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

Dim WshShell, fso
Set WshShell = WScript.CreateObject("WScript.Shell")
Set fso = WScript.CreateObject("Scripting.FileSystemObject")
Dim is_running_windows

is_running_windows = fnIsRunningWindows(param_windir)
If is_running_windows Then
    WScript.Echo ""
    WScript.Echo "***********"
    WScript.Echo "* WARNING *" & "    Seems in a running Windows, DOES NOT ALWAYS WORK."
    WScript.Echo "***********"
    WScript.Echo ""

    If WshShell.Run("net session", 0, True) <> 0 Then
    WScript.Echo ""
    WScript.Echo "***********"
    WScript.Echo "*  ERROR  *" & "    NEED ADMINISTRATIVE(ELEVATED) PRIVILEGES."
    WScript.Echo "***********"
    WScript.Echo ""
    fnExit
    End If
End If

Dim arrSLUI, slui
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
            fnMakeItAccessible full_path
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
            fnMakeItAccessible full_path
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
Function fnIsRunningWindows(full_path)
Dim WshEnv, env_windir, env_username, env_profile
REM USERNAME=SYSTEM
REM USERPROFILE=X:\windows\system32\config\systemprofile

    Set WshEnv   = WshShell.Environment("PROCESS")
    env_windir   = WshEnv("WINDIR")
    env_username = WshEnv("USERNAME")
    env_profile  = WshEnv("USERPROFILE")

    If InStr(LCase(full_path), LCase(env_windir)) > 0 Then
        fnIsRunningWindows = True
    ElseIf InStr(LCase(env_username), LCase("SYSTEM")) > 0 Or InStr(LCase(env_profile), LCase("\windows\system32\config\systemprofile")) > 0 Then
        fnIsRunningWindows = False
    Else
        fnIsRunningWindows = True
    End If
End Function
Function fnMakeItAccessible(ByVal full_path)
    If is_running_windows Then
        WshShell.Run ("TAKEOWN.EXE /F" & " " & """" & full_path & """" & " " & "/A"), 0, True
        WshShell.Run ("ICACLS.EXE" & " " & """" & full_path & """" & " " & "/grant administrators:F"), 0, True
    End If
End Function
