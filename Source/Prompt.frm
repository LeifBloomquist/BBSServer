VERSION 5.00
Begin VB.Form Prompt 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9255
   Icon            =   "Prompt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   9255
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Reply 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   8655
   End
   Begin VB.CommandButton OK 
      Caption         =   "OK"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton Cancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Question"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   8655
   End
End
Attribute VB_Name = "Prompt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public LastResult As String

Public Sub Ask(Q As String, Optional ClearLast = True)
    LastResult = ""
    Label.Caption = Q
    
    If (ClearLast) Then Reply.text = ""
    
    Me.Show vbModal
End Sub


Private Sub Cancel_Click()
    LastResult = CANCELSTRING
    Me.Hide
End Sub

Private Sub OK_Click()
    LastResult = Reply.text
    Me.Hide
End Sub

Private Sub Reply_KeyDown(KeyCode As Integer, Shift As Integer)
    'Enable Enter Key
    If (KeyCode = 13) Then OK_Click
End Sub
