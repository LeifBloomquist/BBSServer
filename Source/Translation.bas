Attribute VB_Name = "Translation"
Option Explicit

'For speed, use two 'mirrored' lookup tables
Dim SerialTranslation(0 To 255) As Byte
Dim TelnetTranslation(0 To 255) As Byte

Public Sub SetupTranslation()

On Error GoTo SetupTranslationError:

    If (TelnetBBS.UsingCharTranslation = False) Then Exit Sub
    
    Dim T As Integer
    
    Dim TextLine As String
    Dim LeftSide As String
    Dim RightSide As String
    Dim EqualSignLoc As Integer
    
    For T = 0 To 255
        SerialTranslation(T) = T
        TelnetTranslation(T) = T
    Next T
    
    Open Advanced.TranslationFile For Input As #1
    
    While (Not EOF(1))
        Line Input #1, TextLine
              
        If Not (Left$(TextLine, 1) = "#") Then  'Ignore Comments
            EqualSignLoc = InStr(TextLine, "=")
            
            If Not (EqualSignLoc = 0) Then      'Ignore lines with no equal sign
                
                LeftSide = Left$(TextLine, EqualSignLoc - 1)
                RightSide = Right$(TextLine, Len(TextLine) - EqualSignLoc)
                
                'Put values into tables
                
                TelnetTranslation(Val(LeftSide)) = Val(RightSide)
                SerialTranslation(Val(RightSide)) = Val(LeftSide)
            End If
        End If
    Wend
    
    Close #1
    
    Exit Sub
    
SetupTranslationError:
    AddMessage "SetupTranslation(): " & Err.Description & " (" & Err.Number & ") " & Advanced.TranslationFile
    TelnetBBS.UsingCharTranslation = False
    Exit Sub
End Sub

Public Function TranslateTelnet(s As String, bytestotal As Long) As String
    Dim T As Integer
    Dim TempString As String
    TempString = ""
    
    For T = 1 To bytestotal
        TempString = TempString & Chr$(TelnetTranslation(Asc(Mid$(s, T, 1))))
    Next T
    
    TranslateTelnet = TempString
End Function

Public Function TranslateSerial(s As String) As String

    Dim T As Integer
    
    Dim TempString As String
    TempString = ""
    
    For T = 1 To Len(s)
        TempString = TempString & Chr$(SerialTranslation(Asc(Mid$(s, T, 1))))
    Next T
    
    TranslateSerial = TempString
End Function
