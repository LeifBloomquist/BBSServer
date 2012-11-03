Attribute VB_Name = "HayesEmulation"
Option Explicit

'Sleep command
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Type VirtualModemType
    '"Real" States
    OffHook As Boolean                  ' The BBS has been taken off-hook temporarily.
    LocalEcho As Boolean                ' Local echo has been set with ATEx
    HardwareFlowControl As Boolean      ' Hardware Flow Control is enabled
    
    'Related to the detection of +++
    OneSecondPauseOccured As Boolean    ' There has been at least one second since the last data sent from the BBS.
    DataReceivedinLastSecond As Boolean ' The BBS has sent data in the last second (used in conjunction with above)
    Pluses As String   'Watch for +++
    
    
    'Special States
    OutGoingCall As Boolean             ' The BBS, or a terminal program, is making an outgoing call.
    WaitingForATA As Boolean            ' Have sent a RING, waiting for ATA.
    
    CommandMode As Boolean              ' The BBS is in command mode after <pause>+++
    ATATimer As Integer                 ' Number of seconds passed while waiting for ATA
    
    'Internal Variables
    LastIPDialled As String             ' Last IP that was called - use with A/
    InvertCase As Boolean               ' Temporary flag, based on case of last command  (Fixes ASCII/PETSCII issues)
    
    ModemCommand As String              ' For AT Commands
End Type

Public VirtualModem As VirtualModemType

Public Sub DoHayesCommand(Command As String)
    Dim ReplyOK As Boolean
    Dim IpAddress As String
    Dim Loc As Integer

On Error GoTo DoHayesCommandError:
    
    'Strip off any extra CR or LF chars
    Command = Replace(Command, Chr$(10), "")
    Command = Replace(Command, Chr$(13), "")
    
    'Check for PETSCII (Commodore) Terminal.  Easy way to do this is to check if command is uppercase.
    
    If (Command = UCase(Command)) Then ' Command was sent as AT.. rather than at.., probably PETSCII
        VirtualModem.InvertCase = True
    Else
        VirtualModem.InvertCase = False
        'Convert to upper case for uniformity
        Command = UCase(Command)
    End If

    If (Advanced.LogHayes.value) Then AddMessage Command
    
    'Strip off "AT".  Carriage Return now removed by HandleCommandMode()
    Command = Mid$(Command, 3, Len(Command) - 2)
    
    'New in 1.2 - Parse through string
    ReplyOK = True
   
'-------------------------------------------------------------------------
'ATD: Dial out.  Must appear all by itself.
'Format is ATDT xx.xx.xx.xx:<port>
    
    If (Left$(Command, 2) = "DT") Or (Left$(Command, 2) = "DP") Then
        
        'Check that calls are allowed
        If (Advanced.AllowOutgoing.value = 0) Then
            SendError
            Exit Sub
        End If
        
        'Get rid of extra spaces in address
        IpAddress = Trim(Right$(Command, Len(Command) - 2))
        
        'Remember IP for A/
        VirtualModem.LastIPDialled = IpAddress
        
        'Make the call!
        MakeOutgoingCall IpAddress
        
        Exit Sub ' So ATD isn't executed below
    End If
    
'-------------------------------------------------------------------------
'ATA:Answer Call. Must appear all by itself.

    If (Left$(Command, 1) = "A") And (VirtualModem.WaitingForATA) Then
        TelnetBBS.ConnectionActive = True
        VirtualModem.WaitingForATA = False
        CallConnected
        Exit Sub
    End If
    
'-------------------------------------------------------------------------
'ATE:Echo

    Loc = InStr(Command, "E")
    If (Loc > 0) Then
        If (Mid$(Command, Loc + 1, 1) = "1") Then
            VirtualModem.LocalEcho = True
        Else
            VirtualModem.LocalEcho = False
        End If
    End If
    
'-------------------------------------------------------------------------
'ATH1, ATH0 or ATH:Hook Commands  (May need +++ first)

   Loc = InStr(Command, "H")
   If (Loc > 0) Then
        Select Case (Mid$(Command & "0", Loc + 1, 1)) 'The & "0" is to prevent running past the string if last command
        
            Case "1":
                TelnetBBS.ShowOffHookState
                VirtualModem.OffHook = True  'Take BBS down until ATH or ATH0 is received
            
            Case Else:
                 GoBackOnline
        End Select
    End If

'-------------------------------------------------------------------------
'ATI:Modem Identification
    If (InStr(Command, "I") > 0) Then
        SerialSendString TelnetBBS.Caption
        SerialSendString "My IP Address is " & IPAddressToUse
    End If
    
'-------------------------------------------------------------------------
'ATD:Go Back Online after +++
    If (InStr(Command, "D") > 0) And (VirtualModem.CommandMode) Then
        If (TelnetBBS.ConnectionActive) Then
            CallConnected
            If (Advanced.LogHayes.value) Then AddMessage "Resumed connection after +++"
        Else
             ReplyOK = False
        End If
    End If
    
'-------------------------------------------------------------------------
'ATU:Shutdown Windows (For headless configurations - must appear by itself)
    If (Left$(Command, 1) = "U") Then
        If (Mid$(Command, 2, 1) = "1") And (Advanced.AllowShutdown.value) Then
            SendOK
            SerialSendString "Shutting down BBS Server and PC."
            ShutdownWindows
            DoEvents
            End
        End If
        
        If (Mid$(Command, 2, 1) = "2") And (Advanced.AllowShutdown.value) Then
            SendOK
            SerialSendString "Shutting down BBS Server."
            DeleteIcon TelnetBBS    'Get rid of the tray icon
            DoEvents
            End              'Exit the program.
        End If
        
        'Shutdown type not given or shutdown not allowed
        SendError
        Exit Sub
    End If
    
'-------------------------------------------------------------------------
'ATZ:Reset Modem (also brute force hangup in some BBS packages)
    If (InStr(Command, "Z") > 0) Then
        ResetVirtualModem True  'True tells it not to send OK, as we do that below
    End If
    
'-------------------------------------------------------------------------
'ATSx Commands
'ATHS0=xxx ' Sets number of rings to answer on.  Ignore the number, but go back on-hook anyway.

   Loc = InStr(Command, "S0=")
   If (Loc > 0) Then GoBackOnline

'-------------------------------------------------------------------------
'  If there were no errors, send OK.  Otherwise send ERROR.

    If (ReplyOK) Then
        SendOK
    Else
        SendError
    End If
    
    Exit Sub
    
DoHayesCommandError:
    AddMessage "DoHayesCommand(): " & Err.Description & " (" & Err.Number & ")"
    Resume Next
End Sub

Public Sub PlusPlusPlusReceived()

    If (Advanced.DisablePlusPlusPlus.value = vbTrue) Then Exit Sub

    If (TelnetBBS.ConnectionActive) Then
        VirtualModem.CommandMode = True
        If (Advanced.LogHayes) Then
            AddMessage "+++"
            AddMessage "Now in Command Mode"
        End If
        
        SendOK
    End If
End Sub

' Send a response string back to the terminal or BBS.
Public Sub SerialSendString(s As String)

    If (Advanced.GuessResponseCase.value = vbUnchecked) Then
        VirtualModem.InvertCase = False
    End If

    If (VirtualModem.InvertCase) Then
        s = Invert(s)
    End If

    If (TelnetBBS.MSComm.PortOpen) Then
        TelnetBBS.MSComm.Output = vbCrLf & s & vbCrLf
        DoEvents
        If (Advanced.LogHayes) Then AddMessage "Responded with " & s
    End If
End Sub

Private Sub SendOK()
    SerialSendString "OK"
End Sub

Private Sub SendError()
    SerialSendString "ERROR"
End Sub

Public Sub CallConnected()
    If (Advanced.SendConnectString) Then
        SerialSendString Advanced.ConnectString
    End If
    
    If (Advanced.LogHayes) Then
        AddMessage "Call Connected"
    End If
    
    VirtualModem.CommandMode = False
End Sub

Public Sub OutgoingCallDisconnected()
    SerialSendString "NO CARRIER"
    VirtualModem.OutGoingCall = False
    VirtualModem.ModemCommand = ""
    DoEvents
End Sub

Public Sub ResetVirtualModem(Optional SuppressOK As Boolean)
    With VirtualModem
        .ATATimer = 0
        .DataReceivedinLastSecond = False
        .OffHook = False
        .OneSecondPauseOccured = False
        .OutGoingCall = False
        .CommandMode = True
        .WaitingForATA = False
        .LastIPDialled = ""
        If (Advanced.EchoCommandChars.value) Then
            .LocalEcho = True
        Else
            .LocalEcho = False
        End If
        
        If (TelnetBBS.ConnectionActive) Then TelnetBBS.Disconnect
        
        TelnetBBS.SerialReset
    End With
    
    If (Not SuppressOK) Then SendOK
End Sub

Public Sub HandleCommandMode(Received As String)

    ' Do Local Echo - Set with ATE0/1
    If (VirtualModem.LocalEcho) Then
        TelnetBBS.SerialTX.Tag = 255
        TelnetBBS.MSComm.Output = Received 'Echo back
        DoEvents
    End If
    
    ' Handle Carriage Return/Linefeed
    If (Right$(Received, 1) = Chr$(13)) Or _
       (Right$(Received, 1) = Chr$(10)) Then
       
       VirtualModem.ModemCommand = VirtualModem.ModemCommand & Received
    
       If (UCase(Left$(VirtualModem.ModemCommand, 2)) = "AT") Then
           DoHayesCommand VirtualModem.ModemCommand
       End If
       
       'Special case, entire command received in 1 'packet' of serial data
       'If (UCase(Left$(Received, 2)) = "AT") Then
       '    DoHayesCommand Received
       'End If
    
       VirtualModem.ModemCommand = ""  'New line, start over again
       Exit Sub
    End If
    
    'Handle Backspace - Either C= DEL or CTRL-H
    If (Right$(Received, 1) = Chr$(8)) Or _
       (Right$(Received, 1) = Chr$(20)) Then
        
        If (Len(VirtualModem.ModemCommand) > 1) Then
            VirtualModem.ModemCommand = Left$(VirtualModem.ModemCommand, Len(VirtualModem.ModemCommand) - 1)
        End If
        
        Exit Sub
    End If
        
    'Append other characters to command string for later processing
    VirtualModem.ModemCommand = VirtualModem.ModemCommand & Received
    
    'Handle A/ - special case
    If UCase(Left$(VirtualModem.ModemCommand, 2)) = "A/" Then
        If (Advanced.LogHayes) Then AddMessage "A/"
        VirtualModem.ModemCommand = ""  'Start over again
        RedialLastIP
        Exit Sub
    End If
End Sub

Public Sub HandlePluses(Received As String)

    ' For Hayes Emulation, check for the <pause>'+++' command.
    ' Everything except the '+' character resets this flag.
    
    If (Right$(Received, 1) = "+") Then
        VirtualModem.Pluses = VirtualModem.Pluses & Received
        VirtualModem.ModemCommand = ""  'Start over again
    Else
        VirtualModem.DataReceivedinLastSecond = True   ' Reset in SecondTimer(), once a second
        VirtualModem.Pluses = ""
    End If
            
    If InStr(VirtualModem.Pluses, "+++") And (VirtualModem.OneSecondPauseOccured) Then
       PlusPlusPlusReceived
       VirtualModem.Pluses = ""   'Start over again
       Exit Sub
    End If
End Sub

Private Sub RedialLastIP()
    If (VirtualModem.LastIPDialled = "") Then
        SendError
    Else
        MakeOutgoingCall VirtualModem.LastIPDialled
    End If
End Sub

Private Function Invert(s As String) As String
    Dim T As Integer
    Dim Char As String
    Invert = ""
    
    For T = 1 To Len(s)
        Char = Mid$(s, T, 1)
        
        If (Char >= "a") And (Char <= "z") Then     'Lowercase
            Invert = Invert & UCase(Char)
        ElseIf (Char >= "A") And (Char <= "Z") Then 'Uppercase
            Invert = Invert & LCase(Char)
        Else                                        'Something else
            Invert = Invert & Char
        End If
    Next T
End Function

Public Sub MakeOutgoingCall(ByVal IpAddress As String)  ' Can't be ByRef, as IpAddress gets modded below

    Dim Port As Long
    Dim ColonPos As Integer
    Dim T As Integer
    
    'Check for Aliases in the Phonebook.  If a match is found, substitute the corresponding Address.
    For T = Advanced.Alias.LBound To Advanced.Alias.UBound
        If LCase(IpAddress) = LCase(Advanced.Alias(T).text) Then
            IpAddress = Advanced.Address(T).text
            Exit For
        End If
    Next T
    
    'Now Continue...

    ColonPos = InStr(IpAddress, ":")
        
    If (ColonPos = 0) Then  'No : found
        Port = 23           'Default
    Else
        'The order of these statements is important!
        Port = Val(Mid$(IpAddress, ColonPos + 1))  'No length=use rest of string
        IpAddress = Left$(IpAddress, ColonPos - 1)
    End If
    
    AddMessage "Outgoing call to " & IpAddress & " Port " & Port
     
    With TelnetBBS
        .ShowOutgoingState
        
        With .Telnet
            If (.State <> sckClosed) Then .Close
            .LocalPort = 0 ' This avoids Winsock run time error 10048.  Reset in TelnetBBS.Disconnect()
            .RemoteHost = IpAddress
            .RemotePort = Port
            .Connect
        End With
        
        .ConnectionActive = True
    End With
    

    VirtualModem.OutGoingCall = True
    
    'Treat just as though it was a BBS caller until disconnect.
    Exit Sub
End Sub

Private Sub GoBackOnline()
    VirtualModem.OffHook = False  'Put BBS Back Up
            
    'If in the middle of a call, hang up.
    If (TelnetBBS.ConnectionActive) And (Advanced.HangupOnATH = vbChecked) Then
        AddMessage "BBS closed connection via ATH or ATS0, closing Telnet session " & TelnetBBS.Telnet.RemoteHostIP
        TelnetBBS.Disconnect
    End If
    
    'Revert back to original state - either waiting or offline
    If TelnetBBS.BoardUp(0).value Then
       TelnetBBS.ShowWaitingState
    Else
       TelnetBBS.ShowOfflineState
    End If
End Sub
