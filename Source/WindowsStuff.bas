Attribute VB_Name = "WindowsStuff"
Option Explicit

'Thanks to Peter Aitken for the PlaySound code, buggy as it was

Public Declare Function PlaySound Lib "winmm.dll" _
  Alias "PlaySoundA" (ByVal lpszName As String, _
  ByVal hModule As Long, ByVal dwFlags As Long) _
  As Long

'The first argument is the name, including the path, of the Wave file to play.
'The second argument isn't used when playing sound files, and you should pass a value of zero. (This function can also play sounds that are associated with system events, but that topic isn't covered here.)
'The final argument consists of flags that control various aspects of how the function works.

'SND_ASYNC (value= 1): play asynchronously, which means that the function returns while the sound is still playing.
'SND_FILENAME (value = &H20000): the first argument is a filename.

Const SND_ASYNC = 1
Const SND_FILENAME = &H20000


'Api functions and the constants required for ExitWindowsEx
Const EWX_LOGOFF = 0
Const EWX_SHUTDOWN = 1
Const EWX_REBOOT = 2
'
' Using this option to shutdown windows does not send
' the WM_QUERYENDSESSION and WM_ENDSESSION messages to
' the open applications. Thus, those apps may loose
' any unsaved data.
'
Const EWX_FORCE = 4
' The following are required to shutdown NT.
'
Const ERROR_NOT_ALL_ASSIGNED = 1300
Const SE_PRIVILEGE_ENABLED = 2
Const TOKEN_QUERY = &H8
Const TOKEN_ADJUST_PRIVILEGES = &H20

Private Type LUID
    lowpart As Long
    highpart As Long
End Type

Private Type LUID_AND_ATTRIBUTES
    pLuid As LUID
    Attributes As Long
End Type

Private Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    Privileges As LUID_AND_ATTRIBUTES
End Type

Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPriv As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As Any, ByVal lpName As String, lpUid As LUID) As Long
Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion      As Long
    dwMinorVersion      As Long
    dwBuildNumber       As Long
    dwPlatformId        As Long
    szCSDVersion        As String * 128
End Type



Private Function FileExists(FullFileName) As Boolean

    ' Passed a filename (with path) returns
    ' True if the file exists, False if not.
    
    Dim s
    
    s = Dir(FullFileName)
    
    If s = "" Then
        FileExists = False
    Else
        FileExists = True
    End If

End Function

Public Sub PlaySoundX(filename As String)

    ' If sound is enabled and filename exists,
    ' play the specified sound.
    
    If FileExists(filename) Then
        PlaySound filename, CLng(0), SND_ASYNC Or SND_FILENAME
    End If
    
End Sub

Public Sub ShutdownWindows()

Dim tLuid          As LUID
Dim tTokenPriv     As TOKEN_PRIVILEGES
Dim tPrevTokenPriv As TOKEN_PRIVILEGES
Dim lResult        As Long
Dim lToken         As Long
Dim lLenBuffer     As Long
Dim lMode As Long
'
' Determine the shutdown mode.
'
' EWX_LOGOFF
'   Shuts down all processes running and
'   logs off the user.
'
' EWX_REBOOT
'   Shuts down and restarts the system.
'
' EWX_SHUTDOWN
'   Shuts down the system to a point where
'   it is safe to turn off the system.
'
' EWX_POWEROFF
'   Shuts down the system and turns off power.
'   The system must support this feature.
'
' EWX_FORCE
'   Forcibly shuts down the system. Files are not closed,...
'   data may be lost.

Dim bWindowsNT As Boolean
'
' Operating System Constants, Types and Declares
'
Const VER_PLATFORM_WIN32s = 0
Const VER_PLATFORM_WIN32_WINDOWS = 1
Const VER_PLATFORM_WIN32_NT = 2

    lMode = EWX_SHUTDOWN
    
    Dim OSInfo As OSVERSIONINFO
    
    OSInfo.dwOSVersionInfoSize = Len(OSInfo)
    Call GetVersionEx(OSInfo)
    bWindowsNT = (OSInfo.dwPlatformId = VER_PLATFORM_WIN32_NT)

    If Not bWindowsNT Then
        Call ExitWindowsEx(lMode, 0)
    Else
        '
        ' Get the access token of the current process.  Get it
        ' with the privileges of querying the access token and
        ' adjusting its privileges.
        '
        lResult = OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, lToken)
        If lResult = 0 Then
            Exit Sub 'Failed
        End If
        '
        ' Get the locally unique identifier (LUID) which
        ' represents the shutdown privilege.
        '
        lResult = LookupPrivilegeValue(0&, "SeShutdownPrivilege", tLuid)
        If lResult = 0 Then Exit Sub 'Failed
        '
        ' Populate the new TOKEN_PRIVILEGES values with the LUID
        ' and allow your current process to shutdown the computer.
        '
        With tTokenPriv
            .PrivilegeCount = 1
            .Privileges.Attributes = SE_PRIVILEGE_ENABLED
            .Privileges.pLuid = tLuid
        lResult = AdjustTokenPrivileges(lToken, False, tTokenPriv, Len(tPrevTokenPriv), tPrevTokenPriv, lLenBuffer)
        End With
        
        If lResult = 0 Then
            Exit Sub 'Failed
        Else
            If Err.LastDllError = ERROR_NOT_ALL_ASSIGNED Then Exit Sub 'Failed
        End If
        '
        '  Shutdown Windows.
        '
        Call ExitWindowsEx(lMode, 0)
    End If
End Sub

