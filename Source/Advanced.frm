VERSION 5.00
Begin VB.Form Advanced 
   BackColor       =   &H00EF7070&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Advanced Settings"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7920
   Icon            =   "Advanced.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   7920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton ShowFrame 
      BackColor       =   &H00EF7070&
      Caption         =   "Connecting"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton ShowFrame 
      BackColor       =   &H00EF7070&
      Caption         =   "Phonebook"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CommandButton ShowFrame 
      BackColor       =   &H00EF7070&
      Caption         =   "Emulation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CommandButton ShowFrame 
      BackColor       =   &H00EF7070&
      Caption         =   "Diagnostics"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton ShowFrame 
      BackColor       =   &H00EF7070&
      Caption         =   "Disconnecting"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton ShowFrame 
      BackColor       =   &H00EF7070&
      Caption         =   "Comms"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton Cancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   6360
      Width           =   1695
   End
   Begin VB.CommandButton ApplyChanges 
      Caption         =   "Save+Apply Changes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   432
      Left            =   120
      TabIndex        =   0
      Top             =   5760
      Width           =   1692
   End
   Begin VB.Frame Frame 
      BackColor       =   &H00EF7070&
      Caption         =   "Diagnostics"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   3495
      Index           =   3
      Left            =   1920
      TabIndex        =   10
      Top             =   120
      Width           =   5895
      Begin VB.CheckBox AllowShutdown 
         BackColor       =   &H00EF7070&
         Caption         =   "Allow Windows Shutdown via ATU"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   44
         ToolTipText     =   "Useful for 'headless' configurations, with no monitor or mouse attached."
         Top             =   3000
         Width           =   3735
      End
      Begin VB.TextBox ConnectionWAV 
         Height          =   285
         Left            =   120
         TabIndex        =   43
         Text            =   "connect.wav"
         Top             =   1680
         Width           =   5055
      End
      Begin VB.CheckBox PlayWAVonConnect 
         BackColor       =   &H00EF7070&
         Caption         =   "Play this WAV file when a caller connects:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   42
         ToolTipText     =   "Lets you know when a caller has connected."
         Top             =   1320
         Value           =   1  'Checked
         Width           =   5712
      End
      Begin VB.TextBox DisconnectWAV 
         Height          =   285
         Left            =   120
         TabIndex        =   41
         Text            =   "disconnect.wav"
         Top             =   2460
         Width           =   5055
      End
      Begin VB.CheckBox PlayWAVonDisconnect 
         BackColor       =   &H00EF7070&
         Caption         =   "Play this WAV file when a caller disconnects:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   40
         ToolTipText     =   "Lets you know when a caller has disconnected."
         Top             =   2100
         Value           =   1  'Checked
         Width           =   5712
      End
      Begin VB.CheckBox LogHayes 
         BackColor       =   &H00EF7070&
         Caption         =   "Detailed Hayes Emulation Logging"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   28
         ToolTipText     =   "Shows all Hayes modem commands received and responses."
         Top             =   600
         Width           =   3735
      End
      Begin VB.CheckBox DetailedDiagnostics 
         BackColor       =   &H00EF7070&
         Caption         =   "Detailed RS-232 Diagnostics Logging"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   "This logs low-level RS-232 events.   (Errors are always logged.)"
         Top             =   240
         Width           =   3975
      End
   End
   Begin VB.Frame Frame 
      BackColor       =   &H00EF7070&
      Caption         =   "Communications"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   2895
      Index           =   2
      Left            =   1920
      TabIndex        =   9
      Top             =   120
      Width           =   5895
      Begin VB.ComboBox SerialSetup 
         Height          =   315
         ItemData        =   "Advanced.frx":0442
         Left            =   120
         List            =   "Advanced.frx":0461
         TabIndex        =   53
         Text            =   "SerialSetup"
         Top             =   480
         Width           =   1875
      End
      Begin VB.CommandButton ShowPinout 
         Caption         =   "Set Defaults"
         Height          =   330
         Left            =   3960
         TabIndex        =   52
         Top             =   1200
         Width           =   1575
      End
      Begin VB.CheckBox EchoCommandChars 
         BackColor       =   &H00EF7070&
         Caption         =   "ATE1 (Local Echo) set by default"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   37
         ToolTipText     =   "This is normally set, but the echoed characters cause problems with Image BBS.  Can be changed at runtime with ATEx."
         Top             =   1980
         Value           =   1  'Checked
         Width           =   4995
      End
      Begin VB.ComboBox CableType 
         Height          =   315
         ItemData        =   "Advanced.frx":04D7
         Left            =   120
         List            =   "Advanced.frx":04EA
         TabIndex        =   35
         Text            =   "CableType"
         ToolTipText     =   "See documentation for pinouts."
         Top             =   1200
         Width           =   3675
      End
      Begin VB.CheckBox EnableFlowControl 
         BackColor       =   &H00EF7070&
         Caption         =   "Enable hardware flow control"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   34
         ToolTipText     =   "Try this when enabling higher baud rates (over 9600 baud)"
         Top             =   1620
         Width           =   3975
      End
      Begin VB.CheckBox EchoTelnetChars 
         BackColor       =   &H00EF7070&
         Caption         =   "Echo characters back to Telnet client"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   19
         ToolTipText     =   "This option should never be needed as most BBS programs will echo on their own."
         Top             =   2340
         Width           =   3975
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "Serial cable type:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   36
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "Serial setup string:  (Baud rate, Parity, Data bits, Stop bits)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   5715
      End
   End
   Begin VB.Frame Frame 
      BackColor       =   &H00EF7070&
      Caption         =   "Connecting"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   4155
      Index           =   0
      Left            =   1920
      TabIndex        =   2
      Top             =   120
      Width           =   5895
      Begin VB.CheckBox RTSOutbound 
         BackColor       =   &H00EF7070&
         Caption         =   "Raise RTS on outbound connection"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   127
         ToolTipText     =   "Only use this if your BBS doesn't natively support PETSCII."
         Top             =   3000
         Width           =   5535
      End
      Begin VB.CheckBox DTROutbound 
         BackColor       =   &H00EF7070&
         Caption         =   "Raise DTR on outbound connection"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   126
         ToolTipText     =   "Only use this if your BBS doesn't natively support PETSCII."
         Top             =   3360
         Width           =   5535
      End
      Begin VB.CheckBox SendRing 
         BackColor       =   &H00EF7070&
         Caption         =   "Send ""RING"" To BBS when Telnet caller connects"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   51
         ToolTipText     =   "If not set, a single carriage return is sent instead."
         Top             =   1560
         Width           =   5652
      End
      Begin VB.CheckBox WaitForATA 
         BackColor       =   &H00EF7070&
         Caption         =   "Wait for BBS to send ATA before completing connection"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   31
         ToolTipText     =   "Will Time Out after 10 seconds.    Also check the Emulation panel for more options."
         Top             =   1200
         Width           =   5715
      End
      Begin VB.CheckBox UseCharTranslation 
         BackColor       =   &H00EF7070&
         Caption         =   "Ask caller for ASCII/PETSCII and use this translation file:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   15
         ToolTipText     =   "Only use this if your BBS doesn't natively support PETSCII."
         Top             =   1920
         Width           =   5535
      End
      Begin VB.TextBox TranslationFile 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   14
         Text            =   "ascii_petscii.ini"
         Top             =   2280
         Width           =   1815
      End
      Begin VB.CheckBox RTSOnConnect 
         BackColor       =   &H00EF7070&
         Caption         =   "Raise RTS when caller connects"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   12
         ToolTipText     =   "This setting depends on your null modem cable and BBS software."
         Top             =   480
         Value           =   1  'Checked
         Width           =   3735
      End
      Begin VB.CheckBox DTROnConnect 
         BackColor       =   &H00EF7070&
         Caption         =   "Raise DTR when caller connects"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   13
         ToolTipText     =   "This setting depends on your null modem cable and BBS software."
         Top             =   840
         Value           =   1  'Checked
         Width           =   3735
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "Incoming Connections"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   128
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "Outbound Connections"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   125
         Top             =   2760
         Width           =   2415
      End
   End
   Begin VB.Frame Frame 
      BackColor       =   &H00EF7070&
      Caption         =   "Disconnecting"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   6615
      Index           =   1
      Left            =   1920
      TabIndex        =   3
      Top             =   120
      Width           =   5895
      Begin VB.CheckBox LowerDTRonDisconnect 
         BackColor       =   &H00EF7070&
         Caption         =   "Lower DTR when caller disconnects"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   63
         ToolTipText     =   "This setting depends on your null modem cable and BBS software."
         Top             =   4560
         Value           =   1  'Checked
         Width           =   3735
      End
      Begin VB.CheckBox LowerRTSonDisconnect 
         BackColor       =   &H00EF7070&
         Caption         =   "Lower RTS when caller disconnects"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   62
         ToolTipText     =   "This setting depends on your null modem cable and BBS software."
         Top             =   4200
         Width           =   3735
      End
      Begin VB.CheckBox SendNoCarrier 
         BackColor       =   &H00EF7070&
         Caption         =   "Send ""NO CARRIER"" To BBS  on disconnect"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   50
         Top             =   3840
         Value           =   1  'Checked
         Width           =   5652
      End
      Begin VB.CheckBox ReEnableRTS 
         BackColor       =   &H00EF7070&
         Caption         =   "Re-Enable RTS after this many seconds"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   720
         TabIndex        =   49
         ToolTipText     =   "Not available when Hardware Flow Control is selected (see Comms)"
         Top             =   5160
         Width           =   4875
      End
      Begin VB.TextBox CarrierDropDelayTime 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         TabIndex        =   48
         Text            =   "5"
         Top             =   5055
         Width           =   495
      End
      Begin VB.CheckBox ReEnableDTR 
         BackColor       =   &H00EF7070&
         Caption         =   "Re-Enable DTR after this many seconds"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   720
         TabIndex        =   47
         ToolTipText     =   "Some BBS's require DTR and RTS only to be low for a few seconds after disconnect."
         Top             =   4875
         Width           =   4875
      End
      Begin VB.CheckBox HangupOnBreak 
         BackColor       =   &H00EF7070&
         Caption         =   "Disconnect if BBS sends RS-232 break"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   38
         ToolTipText     =   "This option is usually *not* needed."
         Top             =   1680
         Width           =   5295
      End
      Begin VB.CheckBox StayOffHook 
         BackColor       =   &H00EF7070&
         Caption         =   "Stay Off-Hook for this many seconds"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   720
         TabIndex        =   30
         ToolTipText     =   "This gives the BBS time to tidy up and reinitialize."
         Top             =   5520
         Width           =   4395
      End
      Begin VB.TextBox StayOffHookTime 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         TabIndex        =   29
         Text            =   "10"
         Top             =   5580
         Width           =   495
      End
      Begin VB.CheckBox HangupOnATH 
         BackColor       =   &H00EF7070&
         Caption         =   "Disconnect if BBS sends <pause>+++ATH<enter>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   20
         ToolTipText     =   "Now working."
         Top             =   1320
         Value           =   1  'Checked
         Width           =   5295
      End
      Begin VB.CheckBox CheckDSR 
         BackColor       =   &H00EF7070&
         Caption         =   "Disconnect if BBS drops DSR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   18
         ToolTipText     =   "This setting depends on your null modem cable and BBS software."
         Top             =   960
         Value           =   1  'Checked
         Width           =   3735
      End
      Begin VB.CheckBox CheckDCD 
         BackColor       =   &H00EF7070&
         Caption         =   "Disconnect if BBS drops DCD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   17
         ToolTipText     =   "This setting depends on your null modem cable and BBS software."
         Top             =   600
         Value           =   1  'Checked
         Width           =   3735
      End
      Begin VB.CheckBox IdleAutoDisconnect 
         BackColor       =   &H00EF7070&
         Caption         =   "Auto-disconnect after this many minutes idle"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   720
         TabIndex        =   8
         Top             =   2400
         Value           =   1  'Checked
         Width           =   4335
      End
      Begin VB.TextBox IdleDisconnectTime 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Text            =   "15"
         Top             =   2445
         Width           =   495
      End
      Begin VB.CheckBox OnlineAutoDisconnect 
         BackColor       =   &H00EF7070&
         Caption         =   "Auto-disconnect after this many minutes online"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   720
         TabIndex        =   6
         Top             =   2040
         Value           =   1  'Checked
         Width           =   4455
      End
      Begin VB.TextBox OnlineDisconnectTime 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Text            =   "180"
         Top             =   2085
         Width           =   495
      End
      Begin VB.CheckBox SendCtrlC 
         BackColor       =   &H00EF7070&
         Caption         =   "Send CHR$(3) (CTRL-C) to BBS on disconnect"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "This option is usually *not* needed."
         Top             =   3480
         Width           =   5055
      End
      Begin VB.Line Line 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   4
         X1              =   825
         X2              =   345
         Y1              =   5715
         Y2              =   5715
      End
      Begin VB.Line Line 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   3
         X1              =   840
         X2              =   360
         Y1              =   5295
         Y2              =   5295
      End
      Begin VB.Line Line 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   2
         X1              =   345
         X2              =   825
         Y1              =   5115
         Y2              =   5115
      End
      Begin VB.Line Line 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   345
         X2              =   825
         Y1              =   2595
         Y2              =   2595
      End
      Begin VB.Line Line 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   0
         X1              =   360
         X2              =   840
         Y1              =   2235
         Y2              =   2235
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "Disconnection Actions"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   46
         Top             =   3240
         Width           =   2415
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "Disconnection Detection"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   45
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.Frame Frame 
      BackColor       =   &H00EF7070&
      Caption         =   "Phone Book / Aliases"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   6615
      Index           =   5
      Left            =   1920
      TabIndex        =   55
      Top             =   120
      Width           =   5895
      Begin VB.CommandButton ImportPB 
         Caption         =   "Import Phone Book"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   432
         Left            =   1320
         TabIndex        =   124
         Top             =   5760
         Width           =   2175
      End
      Begin VB.TextBox Address 
         Height          =   285
         Index           =   19
         Left            =   2400
         TabIndex        =   103
         Top             =   5040
         Width           =   3255
      End
      Begin VB.TextBox Address 
         Height          =   285
         Index           =   18
         Left            =   2400
         TabIndex        =   102
         Top             =   4800
         Width           =   3255
      End
      Begin VB.TextBox Address 
         Height          =   285
         Index           =   17
         Left            =   2400
         TabIndex        =   101
         Top             =   4560
         Width           =   3255
      End
      Begin VB.TextBox Address 
         Height          =   285
         Index           =   16
         Left            =   2400
         TabIndex        =   100
         Top             =   4320
         Width           =   3255
      End
      Begin VB.TextBox Address 
         Height          =   285
         Index           =   15
         Left            =   2400
         TabIndex        =   99
         Top             =   4080
         Width           =   3255
      End
      Begin VB.TextBox Address 
         Height          =   285
         Index           =   14
         Left            =   2400
         TabIndex        =   98
         Top             =   3840
         Width           =   3255
      End
      Begin VB.TextBox Address 
         Height          =   285
         Index           =   13
         Left            =   2400
         TabIndex        =   97
         Top             =   3600
         Width           =   3255
      End
      Begin VB.TextBox Address 
         Height          =   285
         Index           =   12
         Left            =   2400
         TabIndex        =   96
         Top             =   3360
         Width           =   3255
      End
      Begin VB.TextBox Address 
         Height          =   285
         Index           =   11
         Left            =   2400
         TabIndex        =   95
         Top             =   3120
         Width           =   3255
      End
      Begin VB.TextBox Address 
         Height          =   285
         Index           =   10
         Left            =   2400
         TabIndex        =   94
         Top             =   2880
         Width           =   3255
      End
      Begin VB.TextBox Address 
         Height          =   285
         Index           =   9
         Left            =   2400
         TabIndex        =   93
         Top             =   2640
         Width           =   3255
      End
      Begin VB.TextBox Address 
         Height          =   285
         Index           =   8
         Left            =   2400
         TabIndex        =   92
         Top             =   2400
         Width           =   3255
      End
      Begin VB.TextBox Address 
         Height          =   285
         Index           =   7
         Left            =   2400
         TabIndex        =   91
         Top             =   2160
         Width           =   3255
      End
      Begin VB.TextBox Address 
         Height          =   285
         Index           =   6
         Left            =   2400
         TabIndex        =   90
         Top             =   1920
         Width           =   3255
      End
      Begin VB.TextBox Address 
         Height          =   285
         Index           =   5
         Left            =   2400
         TabIndex        =   89
         Top             =   1680
         Width           =   3255
      End
      Begin VB.TextBox Address 
         Height          =   285
         Index           =   4
         Left            =   2400
         TabIndex        =   88
         Top             =   1440
         Width           =   3255
      End
      Begin VB.TextBox Address 
         Height          =   285
         Index           =   3
         Left            =   2400
         TabIndex        =   87
         Top             =   1200
         Width           =   3255
      End
      Begin VB.TextBox Address 
         Height          =   285
         Index           =   2
         Left            =   2400
         TabIndex        =   86
         Top             =   960
         Width           =   3255
      End
      Begin VB.TextBox Address 
         Height          =   285
         Index           =   1
         Left            =   2400
         TabIndex        =   85
         Top             =   720
         Width           =   3255
      End
      Begin VB.TextBox Address 
         Height          =   285
         Index           =   0
         Left            =   2400
         TabIndex        =   84
         Text            =   "bbs.jammingsignal.com:23"
         Top             =   480
         Width           =   3255
      End
      Begin VB.TextBox Alias 
         Height          =   285
         Index           =   19
         Left            =   480
         TabIndex        =   83
         Top             =   5040
         Width           =   1695
      End
      Begin VB.TextBox Alias 
         Height          =   285
         Index           =   18
         Left            =   480
         TabIndex        =   82
         Top             =   4800
         Width           =   1695
      End
      Begin VB.TextBox Alias 
         Height          =   285
         Index           =   17
         Left            =   480
         TabIndex        =   81
         Top             =   4560
         Width           =   1695
      End
      Begin VB.TextBox Alias 
         Height          =   285
         Index           =   16
         Left            =   480
         TabIndex        =   80
         Top             =   4320
         Width           =   1695
      End
      Begin VB.TextBox Alias 
         Height          =   285
         Index           =   15
         Left            =   480
         TabIndex        =   79
         Top             =   4080
         Width           =   1695
      End
      Begin VB.TextBox Alias 
         Height          =   285
         Index           =   14
         Left            =   480
         TabIndex        =   78
         Top             =   3840
         Width           =   1695
      End
      Begin VB.TextBox Alias 
         Height          =   285
         Index           =   13
         Left            =   480
         TabIndex        =   77
         Top             =   3600
         Width           =   1695
      End
      Begin VB.TextBox Alias 
         Height          =   285
         Index           =   12
         Left            =   480
         TabIndex        =   76
         Top             =   3360
         Width           =   1695
      End
      Begin VB.TextBox Alias 
         Height          =   285
         Index           =   11
         Left            =   480
         TabIndex        =   75
         Top             =   3120
         Width           =   1695
      End
      Begin VB.TextBox Alias 
         Height          =   285
         Index           =   10
         Left            =   480
         TabIndex        =   74
         Top             =   2880
         Width           =   1695
      End
      Begin VB.TextBox Alias 
         Height          =   285
         Index           =   9
         Left            =   480
         TabIndex        =   73
         Top             =   2640
         Width           =   1695
      End
      Begin VB.TextBox Alias 
         Height          =   285
         Index           =   8
         Left            =   480
         TabIndex        =   72
         Top             =   2400
         Width           =   1695
      End
      Begin VB.TextBox Alias 
         Height          =   285
         Index           =   7
         Left            =   480
         TabIndex        =   71
         Top             =   2160
         Width           =   1695
      End
      Begin VB.TextBox Alias 
         Height          =   285
         Index           =   6
         Left            =   480
         TabIndex        =   70
         Top             =   1920
         Width           =   1695
      End
      Begin VB.TextBox Alias 
         Height          =   285
         Index           =   5
         Left            =   480
         TabIndex        =   69
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox Alias 
         Height          =   285
         Index           =   4
         Left            =   480
         TabIndex        =   68
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox Alias 
         Height          =   285
         Index           =   3
         Left            =   480
         TabIndex        =   67
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox Alias 
         Height          =   285
         Index           =   2
         Left            =   480
         TabIndex        =   66
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox Alias 
         Height          =   285
         Index           =   1
         Left            =   480
         TabIndex        =   65
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox Alias 
         Height          =   285
         Index           =   0
         Left            =   480
         TabIndex        =   64
         Text            =   "jammingsignal"
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label ID 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "1."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   19
         Left            =   195
         TabIndex        =   123
         Top             =   5130
         Width           =   255
      End
      Begin VB.Label ID 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "1."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   18
         Left            =   195
         TabIndex        =   122
         Top             =   4890
         Width           =   255
      End
      Begin VB.Label ID 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "1."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   17
         Left            =   210
         TabIndex        =   121
         Top             =   4650
         Width           =   255
      End
      Begin VB.Label ID 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "1."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   16
         Left            =   195
         TabIndex        =   120
         Top             =   4395
         Width           =   255
      End
      Begin VB.Label ID 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "1."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   15
         Left            =   195
         TabIndex        =   119
         Top             =   4155
         Width           =   255
      End
      Begin VB.Label ID 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "1."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   14
         Left            =   195
         TabIndex        =   118
         Top             =   3915
         Width           =   255
      End
      Begin VB.Label ID 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "1."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   13
         Left            =   195
         TabIndex        =   117
         Top             =   3660
         Width           =   255
      End
      Begin VB.Label ID 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "1."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   12
         Left            =   195
         TabIndex        =   116
         Top             =   3420
         Width           =   255
      End
      Begin VB.Label ID 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "1."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   11
         Left            =   195
         TabIndex        =   115
         Top             =   3180
         Width           =   255
      End
      Begin VB.Label ID 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "1."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   10
         Left            =   195
         TabIndex        =   114
         Top             =   2940
         Width           =   255
      End
      Begin VB.Label ID 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "1."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   9
         Left            =   195
         TabIndex        =   113
         Top             =   2685
         Width           =   255
      End
      Begin VB.Label ID 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "1."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   8
         Left            =   195
         TabIndex        =   112
         Top             =   2445
         Width           =   255
      End
      Begin VB.Label ID 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "1."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   7
         Left            =   195
         TabIndex        =   111
         Top             =   2205
         Width           =   255
      End
      Begin VB.Label ID 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "1."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   6
         Left            =   195
         TabIndex        =   110
         Top             =   1965
         Width           =   255
      End
      Begin VB.Label ID 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "1."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   195
         TabIndex        =   109
         Top             =   1710
         Width           =   255
      End
      Begin VB.Label ID 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "1."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   195
         TabIndex        =   108
         Top             =   1470
         Width           =   255
      End
      Begin VB.Label ID 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "1."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   195
         TabIndex        =   107
         Top             =   1230
         Width           =   255
      End
      Begin VB.Label ID 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "1."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   195
         TabIndex        =   106
         Top             =   990
         Width           =   255
      End
      Begin VB.Label ID 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "1."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   195
         TabIndex        =   105
         Top             =   735
         Width           =   255
      End
      Begin VB.Label ID 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "1."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   195
         TabIndex        =   104
         Top             =   495
         Width           =   255
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "Address to Call"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   2400
         TabIndex        =   58
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "Alias"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   480
         TabIndex        =   57
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Frame Frame 
      BackColor       =   &H00EF7070&
      Caption         =   "Outgoing calls and Hayes Emulation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   3975
      Index           =   4
      Left            =   1920
      TabIndex        =   25
      Top             =   120
      Width           =   5895
      Begin VB.CheckBox DisablePlusPlusPlus 
         BackColor       =   &H00EF7070&
         Caption         =   "Disable +++ to return to Comand Mode"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   129
         Top             =   3360
         Width           =   4935
      End
      Begin VB.CheckBox GuessResponseCase 
         BackColor       =   &H00EF7070&
         Caption         =   "Try to fix upper/LOWERcase in AT responses"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   61
         ToolTipText     =   "Set this if rEPLIES lOOK lIKE tHIS."
         Top             =   3000
         Width           =   4935
      End
      Begin VB.CheckBox SendWinsockErrors 
         BackColor       =   &H00EF7070&
         Caption         =   "Send Winsock error messages to Terminal program"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   39
         ToolTipText     =   "Useful if you have a ""headless"" system and can't see the logs."
         Top             =   1680
         Value           =   1  'Checked
         Width           =   4935
      End
      Begin VB.TextBox ConnectString 
         Height          =   285
         Left            =   120
         TabIndex        =   33
         Text            =   "CONNECT"
         Top             =   1320
         Width           =   1932
      End
      Begin VB.CheckBox SendConnectString 
         BackColor       =   &H00EF7070&
         Caption         =   "Send this string to BBS when Telnet session connects:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   32
         Top             =   960
         Value           =   1  'Checked
         Width           =   5172
      End
      Begin VB.CheckBox AllowOutgoing 
         BackColor       =   &H00EF7070&
         Caption         =   "Allow Outgoing Calls"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   27
         ToolTipText     =   "Allows BBS or a terminal program to dial out, ATDT xx.xx.xx.xx:<port>"
         Top             =   600
         Value           =   1  'Checked
         Width           =   3975
      End
      Begin VB.CheckBox EnableHayes 
         BackColor       =   &H00EF7070&
         Caption         =   "Enable Hayes Emulation (This keeps COM Port open!)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   2640
         Value           =   1  'Checked
         Width           =   5652
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "Hayes Emulation"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   60
         Top             =   2400
         Width           =   2415
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "Outgoing Calls"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   59
         Top             =   360
         Width           =   2415
      End
   End
End
Attribute VB_Name = "Advanced"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ApplyChanges_Click()
    TelnetBBS.LogDisplay.Clear
    ApplyINI
    SaveINI
    Me.Hide
End Sub

Private Sub Cancel_Click()
    LoadINI  'Loads and re-applies previous settings
    Me.Hide
End Sub

Private Sub ImportPB_Click()
    ImportPhoneBook
End Sub

Private Sub EnableFlowControl_Click()
    RTSOnConnect.value = vbUnchecked
    ReEnableRTS.value = vbUnchecked

    If (EnableFlowControl.value = vbChecked) Then
        RTSOnConnect.Enabled = False
        ReEnableRTS.Enabled = False
        LowerRTSonDisconnect.Enabled = False
        VirtualModem.HardwareFlowControl = True
    End If
    
    If (EnableFlowControl.value = vbUnchecked) Then
        RTSOnConnect.Enabled = True
        ReEnableRTS.Enabled = True
        LowerRTSonDisconnect.Enabled = True
        VirtualModem.HardwareFlowControl = False
    End If
    
End Sub

Private Sub Form_Load()
    ShowFrame_Click (2)
    
    'Set up default IDs in Phonebook
    Dim T
    For T = 0 To 19
        ID(T).Caption = T + 1 & "."
    Next T
End Sub

Private Sub ShowFrame_Click(Index As Integer)
    Dim T As Integer
    
    'Hide all frames and make buttons same color as background
    For T = 0 To Frame.UBound
        Frame(T).Visible = False
        ShowFrame(T).BackColor = Me.BackColor
    Next T
    
    'Show the one that was selected and highlight the button
    Frame(Index).Visible = True
    ShowFrame(Index).BackColor = vbWhite
End Sub

Private Sub ShowPinout_Click()
    Select Case Val(Left$(CableType.text, 1))
    
        Case 1: 'TelBBS Standard Cable
                EnableFlowControl.value = vbChecked
                VirtualModem.HardwareFlowControl = True
                RTSOnConnect.value = vbUnchecked
                DTROnConnect.value = vbChecked
                CheckDCD.value = vbChecked
                CheckDSR.value = vbChecked
                LowerRTSonDisconnect.value = vbUnchecked
                LowerDTRonDisconnect.value = vbChecked

        Case 2: 'Null Modem Cable
                EnableFlowControl.value = vbChecked
                VirtualModem.HardwareFlowControl = True
                RTSOnConnect.value = vbUnchecked
                DTROnConnect.value = vbChecked
                CheckDCD.value = vbChecked
                CheckDSR.value = vbChecked
                LowerRTSonDisconnect.value = vbUnchecked
                LowerDTRonDisconnect.value = vbChecked

        Case 3: 'Radio Shack Cable
                EnableFlowControl.value = vbChecked
                VirtualModem.HardwareFlowControl = True
                RTSOnConnect.value = vbUnchecked
                DTROnConnect.value = vbChecked
                CheckDCD.value = vbChecked
                CheckDSR.value = vbChecked
                LowerRTSonDisconnect.value = vbUnchecked
                LowerDTRonDisconnect.value = vbChecked
                
        Case 4: 'Non-Standard Low Speed Cable
                EnableFlowControl.value = vbUnchecked
                VirtualModem.HardwareFlowControl = False
                RTSOnConnect.value = vbChecked
                DTROnConnect.value = vbChecked
                CheckDCD.value = vbChecked
                CheckDSR.value = vbChecked
                LowerRTSonDisconnect.value = vbChecked
                LowerDTRonDisconnect.value = vbChecked

        Case 5: 'Custom or Unknown Cable
                'Do Nothing
    End Select
End Sub
