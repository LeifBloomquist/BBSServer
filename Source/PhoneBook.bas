Attribute VB_Name = "PhoneBook"
Option Explicit

Public Const CANCELSTRING = "***CANCEL***"
Private Const PHONEBOOKFILE = "phonebook.ini"

Public Sub ImportPhoneBook()
    Dim Filename As String
    
    Prompt.Ask "Please enter complete path to phonebook/configuration file to import:"
    If (Prompt.LastResult = CANCELSTRING) Then Exit Sub
    
    Filename = Prompt.LastResult
    
    LoadPhoneBook Filename
End Sub

'This loads in a phonebook from *either* the tcpser config.xml or phonebook.ini
Public Sub LoadPhoneBook(Optional Filename As String = PHONEBOOKFILE)

On Error GoTo LoadPBError:
    
    Dim temp As String
    Dim valuepos As Integer
    Dim Index As Integer
    Index = 0

    Open Filename For Input As #6
    
    ' Find and extract all entries.
    ' <Entry number="jammingsignal" value="bbs.jammingsignal.com:23" />
    
    While (Not EOF(6))
        Line Input #6, temp
        If InStr(temp, "number=") Then
            'Alias
            Advanced.Alias(Index).text = ExtractQuotes(temp)  'Will extract first string
            
            'Address
            valuepos = InStr(temp, "value=")
            Advanced.Address(Index).text = ExtractQuotes(temp, valuepos) 'Will extract second string
            
            Index = Index + 1
            If (Index >= Advanced.Alias.UBound) Then Exit Sub
        End If
    Wend
    Close #6
    Exit Sub
    
LoadPBError:
    AddMessage "Can't Load Phone Book: " & Err.Description & " (" & Err.Number & ")"
    Exit Sub
End Sub

Public Sub SavePhoneBook()
    'This is not meant to be valid XML, but will try to match the
    'tcpser4j config.xml as closely as possible.
    
On Error Resume Next
    
    Dim T As Integer
    
    Dim Q As String
    Q = Chr$(34)
    
    Open PHONEBOOKFILE For Output As #6
    Print #6, "<PhoneBook>"
    
    ' Output all entries
    ' <Entry number="jammingsignal" value="bbs.jammingsignal.com:23" />
    For T = Advanced.Alias.LBound To Advanced.Alias.UBound
    
        If (Advanced.Alias(T).text = "") Then Exit For
    
        Print #6, Chr$(9) & "<Entry number=" & Q & Advanced.Alias(T).text & Q & _
                            " value=" & Q & Advanced.Address(T).text & Q & " />"
    Next T
         
    Print #6, "</PhoneBook>"
    Close #6
End Sub


'Originally written for gui4cbm4win!
Private Function ExtractQuotes(FullString As String, Optional StartPos As Integer = 1) As String

On Error GoTo QuoteError:

    Dim Quote1 As Integer
    Dim Quote2 As Integer
    
    Quote1 = InStr(StartPos, FullString, Chr$(34))
    Quote2 = InStr(Quote1 + 1, FullString, Chr$(34))
    ExtractQuotes = Mid$(FullString, Quote1 + 1, Quote2 - Quote1 - 1)
    
    Exit Function
    
QuoteError:
     MsgBox "Error: " & Err.Description & " (" & Err.Number & ")  " & Chr$(13) & Chr$(13) & "Debug string: [" & FullString & "] in ExtractQuotes()"
    
End Function
