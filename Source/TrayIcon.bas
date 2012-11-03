Attribute VB_Name = "TrayIcon"
Option Explicit

Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

'Constants
Public Const DOUBLECLICK = 515

Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const NIF_DOALL = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
Public Const WM_MOUSEMOVE = &H200
Public Const NIM_ADD = &H0
Public Const NIM_DELETE = &H2

Private erg As Variant

Public Sub CreateIcon(frm As Form, Title As String)
    Dim Tic As NOTIFYICONDATA
    
On Error GoTo err_CreateIcon:

    Tic.cbSize = Len(Tic)
    Tic.hwnd = frm.hwnd
    Tic.uID = 1&
    Tic.uFlags = NIF_DOALL
    Tic.uCallbackMessage = WM_MOUSEMOVE
    Tic.hIcon = frm.Icon
    Tic.szTip = Title & Chr$(0)
    erg = Shell_NotifyIcon(NIM_ADD, Tic)
    Exit Sub
err_CreateIcon:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
End Sub

Public Sub DeleteIcon(frm As Form)
    Dim Tic As NOTIFYICONDATA
    
On Error GoTo err_DeleteIcon:

    Tic.cbSize = Len(Tic)
    Tic.hwnd = frm.hwnd
    Tic.uID = 1&
    erg = Shell_NotifyIcon(NIM_DELETE, Tic)
    Exit Sub
err_DeleteIcon:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
End Sub

