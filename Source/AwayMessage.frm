VERSION 5.00
Begin VB.Form AwayMessage 
   BackColor       =   &H00EF7070&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Set Maintenance Message"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4680
   ControlBox      =   0   'False
   Icon            =   "AwayMessage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Cancel 
      Caption         =   "Cancel"
      Height          =   312
      Left            =   2760
      TabIndex        =   2
      Top             =   1440
      Width           =   1875
   End
   Begin VB.CommandButton OK 
      Caption         =   "OK (Save Message)"
      Height          =   312
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   1875
   End
   Begin VB.TextBox Message 
      Height          =   1215
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "AwayMessage.frx":000C
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "AwayMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Cancel_Click()
    LoadAwayMessage  'Load previous message
    Me.Hide
End Sub

Private Sub OK_Click()
    Open "c64bbsaway.ini" For Output As #2
    Print #2, AwayMessage.Message.text
    Close #2
    
    Me.Hide
End Sub

Public Sub LoadAwayMessage()

On Error GoTo LoadMessageError:

    Dim temp As String
    
    Open "c64bbsaway.ini" For Input As #2
    
    AwayMessage.Message = ""
    
    While Not EOF(2)
        Line Input #2, temp
        AwayMessage.Message = AwayMessage.Message & temp
    Wend
    
    Close #2
    Exit Sub
    
LoadMessageError:

    If Err.Number = 53 Then  'File not found
        OK_Click ' Force save
    Else
        AddMessage " LoadAwayMessage(): " & Err.Description & " (" & Err.Number & ")"
    End If
    
    Exit Sub

End Sub
