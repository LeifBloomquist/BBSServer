Attribute VB_Name = "INIRoutines"
Option Explicit

Const INIFILE = "telnetbbs.ini"

'Apply INI file contents
Public Sub ApplyINI()

On Error GoTo ApplyINIError:

    With TelnetBBS
        .SerialReset
        
        'Load in the away message
        AwayMessage.LoadAwayMessage
        
        .IPAddressList.text = IPAddressToUse
        
        .Incoming(0).Close
        .Incoming(0).Protocol = sckTCPProtocol                       'Must be TCP for Telnet
        .Incoming(0).LocalPort = .TelnetPortText.text                'Specifies port to listen on
        .Incoming(0).Bind Val(.TelnetPortText.text), IPAddressToUse  'Binds to a specific NIC
        .Incoming(0).Listen                                          'Put into Listen mode
        
        'Load+Set up character translation, if active
        SetupTranslation
    
        AddMessage TelnetBBS.Caption & " listening on IP " & .IPAddressList.text & ", Port " & .Incoming(0).LocalPort
    End With
        
    Exit Sub
    
ApplyINIError:
    AddMessage "ApplyINI(): " & Err.Description & " (" & Err.Number & ")"
    Exit Sub
End Sub


'Load the INI file
Public Sub LoadINI()
On Error GoTo LoadINIError  'Get through as much as possible
    Dim ErrorFound As Boolean
    ErrorFound = False

    'Load Phonebook first, even if there's no config file
    LoadPhoneBook
    
    'Load the parameters from the ini file
    Open INIFILE For Input As #1
    
    '[BBS]
    TelnetBBS.BBSNameText.text = GetINIValue("BBSName")
    TelnetBBS.TelnetPortText.text = GetINIValue("TelnetPort")
    IPAddressToUse = GetINIValue("IPAddress")
    TelnetBBS.COMportText.text = GetINIValue("COMPort")
   
    '[Connecting]
    Advanced.RTSOnConnect.value = GetINIValue("RTSOnConnect")
    Advanced.DTROnConnect.value = GetINIValue("DTROnConnect")
    Advanced.WaitForATA.value = GetINIValue("WaitForATA")
    Advanced.UseCharTranslation.value = GetINIValue("UseCharTranslation")
    Advanced.TranslationFile.text = GetINIValue("TranslationFile")
    Advanced.RTSOutbound.value = GetINIValue("RTSOutbound")
    Advanced.DTROutbound.value = GetINIValue("DTROutbound")
    
    '[Disconnecting]
    Advanced.SendCtrlC.value = GetINIValue("SendCtrlC")
    Advanced.CheckDCD.value = GetINIValue("CheckDCD")
    Advanced.CheckDSR.value = GetINIValue("CheckDSR")
    Advanced.HangupOnATH.value = GetINIValue("HangupOnATH")
    Advanced.OnlineAutoDisconnect.value = GetINIValue("OnlineAutoDisconnect")
    Advanced.OnlineDisconnectTime.text = GetINIValue("OnlineDisconnectTime")
    Advanced.IdleAutoDisconnect.value = GetINIValue("IdleAutoDisconnect")
    Advanced.IdleDisconnectTime.text = GetINIValue("IdleDisconnectTime")
    Advanced.LowerRTSonDisconnect.value = GetINIValue("LowerRTSonDisconnect")
    Advanced.LowerDTRonDisconnect.value = GetINIValue("LowerDTRonDisconnect")
    
    Advanced.ReEnableDTR.value = GetINIValue("ReEnableDTR")
    Advanced.ReEnableRTS.value = GetINIValue("ReEnableRTS")
    Advanced.CarrierDropDelayTime.text = GetINIValue("CarrierDropDelayTime")
    Advanced.StayOffHook.value = GetINIValue("StayOffHook")
    Advanced.StayOffHookTime.text = GetINIValue("StayOffHookTime")
    Advanced.HangupOnBreak = GetINIValue("HangupOnBreak")

    '[Comms]
    Advanced.SerialSetup.text = GetINIValue("SerialSetup")
    Advanced.EchoTelnetChars.value = GetINIValue("EchoTelnetChars")
    Advanced.EnableFlowControl.value = GetINIValue("EnableFlowControl")
    Advanced.EchoCommandChars.value = GetINIValue("EchoCommandChars")
    Advanced.CableType.text = GetINIValue("CableType")

    '[Diagnostics]
    Advanced.DetailedDiagnostics.value = GetINIValue("DetailedDiagnostics")
    Advanced.LogHayes.value = GetINIValue("LogHayes")
    Advanced.PlayWAVonConnect.value = GetINIValue("PlayWAVonConnect")
    Advanced.ConnectionWAV.text = GetINIValue("ConnectionWAV")
    Advanced.PlayWAVonDisconnect.value = GetINIValue("PlayWAVonDisconnect")
    Advanced.DisconnectWAV.text = GetINIValue("DisconnectWAV")
    Advanced.AllowShutdown = GetINIValue("AllowShutdown")

    '[Emulation]
    Advanced.EnableHayes.value = GetINIValue("EnableHayes")
    Advanced.AllowOutgoing.value = GetINIValue("AllowOutgoing")
    Advanced.SendRing.value = GetINIValue("SendRing")
    Advanced.SendConnectString.value = GetINIValue("SendConnectString")
    Advanced.ConnectString.text = GetINIValue("ConnectString")
    Advanced.SendNoCarrier.value = GetINIValue("SendNoCarrier")
    Advanced.SendWinsockErrors.value = GetINIValue("SendWinsockErrors")
    Advanced.GuessResponseCase.value = GetINIValue("GuessResponseCase")
    Advanced.DisablePlusPlusPlus.value = GetINIValue("DisablePlusPlusPlus")
    
    '[EOF]
    Close #1
    
    If (ErrorFound) Then MsgBox "This appears to be the first time you have run this version.  Please check your options.", vbOKOnly, "First Run"
    Exit Sub
    
LoadINIError:
    ErrorFound = True
    Resume Next
    
    'Set defaults on the drop-down menus, as they can't be overridden
    'Advanced.SerialSetup.text = Advanced.SerialSetup.List(2)  '2400 bps
    'Advanced.CableType.text = Advanced.CableType.List(0)      'Standard Cable
End Sub



'Write the INI file
Public Sub SaveINI()

On Error GoTo SaveINIError:

    'Save the values into the ini file
    Open INIFILE For Output As #1
    
    Print #1, "#Version = " & TelnetBBS.Caption
    Print #1, ""
    
    Print #1, "[BBS]"
    PutINIValue "BBSName", TelnetBBS.BBSNameText.text
    PutINIValue "TelnetPort", TelnetBBS.TelnetPortText.text
    PutINIValue "IPAddress", TelnetBBS.IPAddressList.text
    PutINIValue "COMPort", TelnetBBS.COMportText.text
    
    Print #1, ""
    Print #1, "[Connecting]"
    PutINIValue "RTSOnConnect", Advanced.RTSOnConnect.value
    PutINIValue "DTROnConnect", Advanced.DTROnConnect.value
    PutINIValue "WaitForATA", Advanced.WaitForATA.value
    PutINIValue "UseCharTranslation", Advanced.UseCharTranslation.value
    PutINIValue "TranslationFile", Advanced.TranslationFile.text
    PutINIValue "RTSOutbound", Advanced.RTSOutbound.value
    PutINIValue "DTROutbound", Advanced.DTROutbound.value
    
    Print #1, ""
    Print #1, "[Disconnecting]"
    PutINIValue "SendCtrlC", Advanced.SendCtrlC.value
    PutINIValue "CheckDCD", Advanced.CheckDCD.value
    PutINIValue "CheckDSR", Advanced.CheckDSR.value
    PutINIValue "HangupOnATH", Advanced.HangupOnATH.value
    PutINIValue "OnlineAutoDisconnect", Advanced.OnlineAutoDisconnect.value
    PutINIValue "OnlineDisconnectTime", Advanced.OnlineDisconnectTime.text
    PutINIValue "IdleAutoDisconnect", Advanced.IdleAutoDisconnect.value
    PutINIValue "IdleDisconnectTime", Advanced.IdleDisconnectTime.text
    PutINIValue "LowerRTSonDisconnect", Advanced.LowerRTSonDisconnect.value
    PutINIValue "LowerDTRonDisconnect", Advanced.LowerDTRonDisconnect.value
    PutINIValue "ReEnableDTR", Advanced.ReEnableDTR.value
    PutINIValue "ReEnableRTS", Advanced.ReEnableRTS.value
    PutINIValue "CarrierDropDelayTime", Advanced.CarrierDropDelayTime.text
    PutINIValue "StayOffHook", Advanced.StayOffHook.value
    PutINIValue "StayOffHookTime", Advanced.StayOffHookTime.text
    PutINIValue "HangupOnBreak", Advanced.HangupOnBreak.value
    
    Print #1, ""
    Print #1, "[Comms]"
    PutINIValue "SerialSetup", Advanced.SerialSetup.text
    PutINIValue "EchoTelnetChars", Advanced.EchoTelnetChars.value
    PutINIValue "EnableFlowControl", Advanced.EnableFlowControl.value
    PutINIValue "EchoCommandChars", Advanced.EchoCommandChars.value
    PutINIValue "CableType", Advanced.CableType.text
    
    Print #1, ""
    Print #1, "[Diagnostics]"
    PutINIValue "DetailedDiagnostics", Advanced.DetailedDiagnostics.value
    PutINIValue "LogHayes", Advanced.LogHayes.value
    PutINIValue "PlayWAVonConnect", Advanced.PlayWAVonConnect.value
    PutINIValue "ConnectionWAV", Advanced.ConnectionWAV.text
    PutINIValue "PlayWAVonDisconnect", Advanced.PlayWAVonDisconnect.value
    PutINIValue "DisconnectWAV", Advanced.DisconnectWAV.text
    PutINIValue "AllowShutdown", Advanced.AllowShutdown.value

    Print #1, ""
    Print #1, "[Emulation]"
    PutINIValue "EnableHayes", Advanced.EnableHayes.value
    PutINIValue "AllowOutgoing", Advanced.AllowOutgoing.value
    PutINIValue "SendRing", Advanced.SendRing.value
    PutINIValue "SendConnectString", Advanced.SendConnectString.value
    PutINIValue "ConnectString", Advanced.ConnectString.text
    PutINIValue "SendNoCarrier", Advanced.SendNoCarrier.value
    PutINIValue "SendWinsockErrors", Advanced.SendWinsockErrors.value
    PutINIValue "GuessResponseCase", Advanced.GuessResponseCase.value
    PutINIValue "DisablePlusPlusPlus", Advanced.DisablePlusPlusPlus.value

    Print #1, ""
    Print #1, "[EOF]"
    Close #1
    
    MsgBox "Settings applied and saved to " & Chr$(13) & Chr$(13) & CurDir & "\" & INIFILE, vbInformation, "BBS Server"
    
    'Save Phonebook as well
    SavePhoneBook
    
    Exit Sub

SaveINIError:
    Close #1
    AddMessage "SaveINI(): " & Err.Description & " (" & Err.Number & ")"
    Exit Sub
    
End Sub

Private Function GetINIValue(valname As String) As Variant

On Error GoTo GetINIValueError:

    Dim temp As String
    Dim EqualLoc As Integer
    Seek #1, 1
  
    While Not EOF(1)
        Line Input #1, temp
        If Left$(temp, Len(valname)) = valname Then
            EqualLoc = InStr(temp, "=")
            
            If (EqualLoc = 0) Then
                AddMessage "Error: Corrupt line in INI file"
                AddMessage temp
                Exit Function
             Else
                GetINIValue = Mid$(temp, EqualLoc + 1)
                Exit Function
            End If
        End If
    Wend
    
    AddMessage "Error: INI file missing entry for " & valname
    GetINIValue = Null
    Exit Function
    
GetINIValueError:
    Close #1
    AddMessage "GetINIValueError(): " & Err.Description & " (" & Err.Number & ")"
    Exit Function
End Function

Private Sub PutINIValue(valname As String, value As Variant)
    Print #1, valname & "=" & CStr(value)
End Sub

Public Sub AddMessage(Message As String)
    TelnetBBS.LogDisplay.AddItem Now & "  " & Message, 0
    
    Open "c64bbslog.txt" For Append As #10
    Print #10, Now & " " & Message
    Close #10
    
    If (TelnetBBS.LogDisplay.ListCount >= 1000) Then
        TelnetBBS.LogDisplay.Clear
        TelnetBBS.LogDisplay.AddItem "Auto-cleared after 1000 messages."
    End If
End Sub
