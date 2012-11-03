Attribute VB_Name = "Telnet"
Option Explicit

Public Function IACFilter(sInput As String) As String
    Do Until InStr(sInput, Chr(250)) = 0 'Loop and remove sub IAC negotiations first
        sInput = Replace(sInput, Mid$(sInput, InStr(sInput, Chr(250)), InStr(sInput, Chr(240)) - InStr(sInput, Chr(250))), "")
    Loop
    
    Do Until InStr(sInput, Chr(255)) = 0 'Loop and remove normal IAC commands
        sInput = Replace(sInput, Mid$(sInput, InStr(sInput, Chr(255)), InStr(sInput, Chr(255)) + 3), "")
    Loop 'Until InStr(sInput, Chr(255)) = 0
    
    IACFilter = sInput
End Function

Public Function IACResponse(sInput As String) As String
'pass winsock1.getdata string to this function
'this function will compose a rejection response to all the server requests

    Dim iPos As Integer
    Dim iRet As Integer
    
    IACResponse = "" 'reset string
    iPos = 1 'reset counter
    
    Do Until InStr(iPos, sInput, Chr(255)) = 0 'loop until no more IAC's found
        iPos = InStr(iPos, sInput, Chr(255)) 'set iPos if finds 1st or next IAC command(255)
        'respond to command
        Select Case Mid$(sInput, iPos + 1, 1)
            Case Chr(254) 'DONT respond with a WONT
                 IACResponse = IACResponse & Chr(255) & Chr(252) & Mid$(sInput, iPos + 2, 1)
                 
            Case Chr(253) 'DO respond with a WONT
                 IACResponse = IACResponse & Chr(255) & Chr(252) & Mid$(sInput, iPos + 2, 1)
                 
            Case Chr(252) 'WONT do not respond!
                 IACResponse = IACResponse
                 
            Case Chr(251) 'WILL respond with a DONT
                 IACResponse = IACResponse & Chr(255) & Chr(254) & Mid$(sInput, iPos + 2, 1)
                
            Case Chr(250) 'Beginning of sub negotiation
                 IACResponse = IACResponse 'do not respond skip to end of 240
                 iPos = InStr(iPos, sInput, Chr(240))
                 
            Case Else
                 AddMessage "Cannot respond to IAC negotiation command " & Format(Asc(Mid$(sInput, iPos + 1, 1)), "000")
        End Select
        
        iPos = iPos + 1
    Loop
    
    'function should return either "" or a rejection IACResponse
End Function
