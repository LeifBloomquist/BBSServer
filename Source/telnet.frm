VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form TelnetBBS 
   BackColor       =   &H00EF7070&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Telnet BBS Server v1.4a"
   ClientHeight    =   8520
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   7410
   Icon            =   "telnet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   568
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   494
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame 
      BackColor       =   &H00EF7070&
      Caption         =   "BBS Status"
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
      Height          =   1395
      Index           =   5
      Left            =   1800
      TabIndex        =   39
      Top             =   0
      Width           =   2235
      Begin VB.CommandButton ForceDisconnect 
         Caption         =   "Force Disconnect"
         Height          =   1065
         Left            =   120
         Picture         =   "telnet.frx":34CA
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   240
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.OptionButton BoardUp 
         BackColor       =   &H00EF7070&
         Caption         =   "Accepting Calls"
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
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   42
         Top             =   240
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton BoardUp 
         BackColor       =   &H00EF7070&
         Caption         =   "Not Accepting Calls"
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
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   41
         Top             =   600
         Width           =   2055
      End
      Begin VB.CommandButton SetAwayMessage 
         Caption         =   "Set Message..."
         Height          =   330
         Left            =   120
         TabIndex        =   40
         Top             =   960
         Width           =   2055
      End
   End
   Begin VB.Frame Frame 
      BackColor       =   &H00EF7070&
      Caption         =   "Telnet Activity"
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
      Height          =   1395
      Index           =   3
      Left            =   120
      TabIndex        =   34
      Top             =   0
      Width           =   1575
      Begin VB.Timer OffHookTimer 
         Enabled         =   0   'False
         Left            =   1080
         Top             =   240
      End
      Begin VB.CommandButton TelnetRX 
         BackColor       =   &H00000000&
         Enabled         =   0   'False
         Height          =   195
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   1035
         Width           =   255
      End
      Begin VB.CommandButton TelnetTX 
         BackColor       =   &H00000000&
         Enabled         =   0   'False
         Height          =   195
         Left            =   855
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   1035
         Width           =   255
      End
      Begin VB.Label BBSState 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   165
         TabIndex        =   43
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "RX"
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
         Index           =   18
         Left            =   330
         TabIndex        =   38
         Top             =   795
         Width           =   315
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "TX"
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
         Index           =   17
         Left            =   840
         TabIndex        =   37
         Top             =   795
         Width           =   315
      End
   End
   Begin VB.Frame Frame 
      BackColor       =   &H00EF7070&
      Caption         =   "Serial Activity"
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
      Height          =   795
      Index           =   4
      Left            =   120
      TabIndex        =   19
      ToolTipText     =   "These LEDs show the state of the signals on the PC side."
      Top             =   1440
      Width           =   3915
      Begin VB.Timer GraphicsTimer 
         Interval        =   30
         Left            =   -120
         Top             =   240
      End
      Begin VB.CommandButton SerialDTR 
         BackColor       =   &H00000000&
         Height          =   195
         Left            =   3300
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   480
         Width           =   255
      End
      Begin VB.CommandButton SerialRTS 
         BackColor       =   &H00000000&
         Height          =   195
         Left            =   2805
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   480
         Width           =   255
      End
      Begin VB.CommandButton SerialDCD 
         BackColor       =   &H00000000&
         Enabled         =   0   'False
         Height          =   195
         Left            =   2325
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   480
         Width           =   255
      End
      Begin VB.CommandButton SerialDSR 
         BackColor       =   &H00000000&
         Enabled         =   0   'False
         Height          =   195
         Left            =   1830
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   480
         Width           =   255
      End
      Begin VB.CommandButton SerialCTS 
         BackColor       =   &H00000000&
         Enabled         =   0   'False
         Height          =   195
         Left            =   1335
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   480
         Width           =   255
      End
      Begin VB.CommandButton SerialTX 
         BackColor       =   &H00000000&
         Enabled         =   0   'False
         Height          =   195
         Left            =   855
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   480
         Width           =   255
      End
      Begin VB.CommandButton SerialRX 
         BackColor       =   &H00000000&
         Enabled         =   0   'False
         Height          =   195
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "DTR"
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
         Index           =   11
         Left            =   3240
         TabIndex        =   33
         Top             =   240
         Width           =   435
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "RTS"
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
         Index           =   10
         Left            =   2745
         TabIndex        =   31
         Top             =   240
         Width           =   435
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "DCD"
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
         Index           =   9
         Left            =   2250
         TabIndex        =   29
         Top             =   240
         Width           =   435
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "DSR"
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
         Index           =   8
         Left            =   1740
         TabIndex        =   27
         Top             =   240
         Width           =   435
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "CTS"
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
         Index           =   7
         Left            =   1275
         TabIndex        =   22
         Top             =   240
         Width           =   435
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "TX"
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
         Index           =   6
         Left            =   840
         TabIndex        =   21
         Top             =   240
         Width           =   315
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "RX"
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
         Left            =   330
         TabIndex        =   20
         Top             =   240
         Width           =   315
      End
   End
   Begin VB.CommandButton Exit 
      BackColor       =   &H000000FF&
      Caption         =   "Exit"
      Height          =   312
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   8160
      Width           =   975
   End
   Begin VB.Timer SecondTimer 
      Interval        =   1000
      Left            =   6840
      Top             =   7560
   End
   Begin VB.Frame Frame 
      BackColor       =   &H00EF7070&
      Caption         =   "Caller Status"
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
      Height          =   795
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   2280
      Width           =   3915
      Begin VB.Label CallerStatus 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "..."
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
         Left            =   120
         TabIndex        =   15
         Top             =   345
         Width           =   3675
      End
   End
   Begin VB.Frame Frame 
      BackColor       =   &H00EF7070&
      Caption         =   "Configuration"
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
      Height          =   3075
      Index           =   1
      Left            =   4140
      TabIndex        =   3
      Top             =   0
      Width           =   3195
      Begin VB.ComboBox IPAddressList 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   900
         Width           =   1695
      End
      Begin VB.CommandButton AdvancedButton 
         Caption         =   "Advanced..."
         Height          =   312
         Left            =   120
         TabIndex        =   16
         Top             =   2190
         Width           =   2895
      End
      Begin VB.TextBox BBSNameText 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Text            =   "Commodore Telnet BBS"
         Top             =   480
         Width           =   2895
      End
      Begin VB.TextBox COMportText 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         TabIndex        =   6
         Text            =   "1"
         Top             =   1770
         Width           =   1695
      End
      Begin VB.TextBox TelnetPortText 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         TabIndex        =   5
         Text            =   "23"
         Top             =   1350
         Width           =   1695
      End
      Begin VB.CommandButton ApplyChanges 
         Caption         =   "Save+Apply Changes"
         Height          =   312
         Left            =   120
         TabIndex        =   4
         Top             =   2640
         Width           =   2895
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "BBS Name / Welcome Banner:"
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
         Height          =   252
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   2952
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "COM Port:"
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
         Index           =   3
         Left            =   120
         TabIndex        =   9
         Top             =   1800
         Width           =   1275
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "Telnet Port:"
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
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   1334
         Width           =   1275
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "IP Address:"
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
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   922
         Width           =   1275
      End
   End
   Begin MSWinsockLib.Winsock Incoming 
      Index           =   0
      Left            =   3900
      Top             =   6480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   1000
   End
   Begin MSWinsockLib.Winsock Telnet 
      Left            =   4320
      Top             =   6480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   1000
   End
   Begin MSCommLib.MSComm MSComm 
      Left            =   4800
      Top             =   6480
      _ExtentX        =   794
      _ExtentY        =   794
      _Version        =   393216
      CommPort        =   9
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton ClearOnly 
      Caption         =   "Clear Log"
      Height          =   312
      Left            =   120
      TabIndex        =   0
      Top             =   8160
      Width           =   2175
   End
   Begin VB.CommandButton About 
      Caption         =   "About..."
      Height          =   312
      Left            =   5340
      TabIndex        =   1
      Top             =   8160
      Width           =   975
   End
   Begin VB.Frame Frame 
      BackColor       =   &H00EF7070&
      Caption         =   "Activity Log (Newest Events First)"
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
      Height          =   4935
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Top             =   3120
      Width           =   7215
      Begin VB.ListBox LogDisplay 
         Height          =   4155
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   6975
      End
      Begin VB.Label OutCount 
         BackStyle       =   0  'Transparent
         Caption         =   "-"
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
         Left            =   120
         TabIndex        =   44
         Top             =   4560
         Width           =   5475
      End
   End
   Begin VB.Image WaitingIcon 
      Height          =   480
      Left            =   0
      Picture         =   "telnet.frx":37D4
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image OfflineIcon 
      Height          =   480
      Left            =   0
      Picture         =   "telnet.frx":6C9E
      Top             =   1200
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image OnlineIcon 
      Height          =   480
      Left            =   0
      Picture         =   "telnet.frx":A168
      Top             =   720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Log file is c64bbslog.txt"
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   2340
      TabIndex        =   2
      Top             =   8220
      Width           =   1992
   End
End
Attribute VB_Name = "TelnetBBS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Commodore 64 (etc...) Telnet BBS program.
' Bridges data between a TCP port and a COM port.
'
' Copyright 2003-2006 Leif Bloomquist leifb@ica.net   http://home.ica.net/~leifb/bbs/
'
' You may freely use, modify, distribute etc. this source code, but please keep
' my copyright notice in both the source and the "About" box.
'
' No warranty of any kind.

' Version 0.8a - General release.  (LB)
' Version 0.8b - Added detection for DSR *or* DCD to be dropped.  (LB)
' Version 0.8c - Replaced 'Sleep()' with 'DoEvents' after down/busy messages.  (LB)
' Version 0.8d - Added option to set DTR as well as RTS on connect,
'                expanded diagnostics and Advanced page. Added Sounds. (LB)
' Version 0.9a - Finally - Hayes Emulation.  Also included Fox's changes.  Revamped config file.
'         0.9b - Bugfixes
'
' Version 1.0a - RC1
'                Mega Public release.
'                Added IAC Telnet Response Code.
'                Fixed bug with COM port initialization order.
'                Fixed bug with ATH0 wrongly saying BBS was in "Waiting" state.
'                CR or LF can now be used as delimiter before AT Commands.
'
'                RC2
'                Got rid of the popup boxes if an unknown IAC code is received (when a Real Telnet session is in progress).
'                Check the first incoming byte of the session, and 255="Real" Telnet session (i.e. to Linux), anything else means a Telnet BBS.
'                       In the latter case, data will be passed transparently so file transfers should work.
'                If there is only 1 IP address on the system, always use it.
'                Replaced the Sleep() for staying offhook with a Timer.
'                Added an option for echoing data back to BBS when in command mode.
'
'                RC3 (final)
'                Changed Connecting... to connecting...
'                Fixed bug where "NO CARRIER" was sent even if nothing connected, which caused a RunTime error.
'                Added option to hang up on an RS-232 break.
'                Added option for sending Winsock errors to Terminal program.
'                General code tidyup.
'
' Version 1.1
'                Bugfixes and tidyup.
'                Enabled backspace when typing an AT command. (finally!)
'                Added auto-detection of command case (i.e. PETSCII), sO rESPONSES aREN'T lIKE tHIS
'                Added A/ support.
'                Added ATEx support.
'                Switched to a 'virtual modem' internal state structure.
'                ATZ resets the virtual modem.
'                Option to play a WAV file when callers connect or disconnect
'                ATU (non-standard command) shuts down Windows, for headless systems. (!)
'                  ATU1 shuts down PC, ATU2 just exits BBS Server
'                Pressing a key aborts the current call before it connects.
'
' Version 1.2
'                Added an option to enable hardware flow control.
'                Fancier Hayes emulation to parse entire string.
'                Redesigned GUI to show more serial activity detail.
'                Fixed A/ bug.
'                Sends a busy file (busy.txt) to callers when busy
'                Added Phone Book and import routine.
'                Better control of signals at disconnect.
'                Different cable types and defaults.
'
' Version 1.3b1  Fixed a bug with repeated file #s.
'                Fixed bug that truncated AT commands at 1200 baud (i.e calling Qlink)
'                Changed MSComm.OutBufferSize to 10000
'                Better error handling in phone book load
'                Added a force disconnect button
'
' Version 1.3    (Fixes by Eric Pearson)
'                Raise RTS on outbound connection
'                Raise DTR on outbound connection
'                Lower RTS when Call disconnects and Lower DTR when caller disconnects are preserved in INI
'
' Version 1.4    Option to force RTS always on by adding -rts to the command line.
'                Fixed bug where A/ only remembered the port a single time.
'                Option to disable +++ detection.
'                ATS0 will put BBS back on-hook.
'
' Version 1.4a   Bugfix!  Disconnect on RS232 Break only worked in Diag mode before.
'                Bugfix!  Port for Outgoing calls is now a Long (was Integer)
'
'TODO:
'                Phone book bug!   Book is saved when form not loaded
'                Numerical response codes (needed for some BBSes i.e. Image)
'                Weird bug where file transfers fail after 255 blocks - how would BBS Server cause that though?

Option Explicit

'Global Variables
Public ConnectionActive As Boolean      ' This is set when a 'caller' is connected.
Dim ConnectionTime As Long              ' This counts the number of seconds caller has been connected
Dim IdleTime As Long                    ' This counts the number of seconds caller has been idle
Dim WaitingForResponse As Boolean       ' This is set when waiting for a special character from the caller
Public UsingCharTranslation As Boolean  ' This is set when the caller wants translation, i.e. PETSCII mode
Dim ProgramShutDown As Boolean          ' Internal flag used when shutting down, to keep things tidy (i.e. tray icon)
Public FirstCharReceived As Boolean     ' This flag is set after the first character is received.
Public TrueTelnetMode As Boolean        ' Set if the first character was byte 255 (Telnet IAC)
Public SerialBufCountMax As Integer     ' Maximum buffer count for this call
Public ForceRTS As Boolean              ' Hack for some BBSes (i.e. Mad World)

Private Sub About_Click()
    MsgBox TelnetBBS.Caption & Chr$(13) & _
           "Copyright 2003-2008 by Leif Bloomquist" & Chr$(13) & Chr$(13) & _
           "For more information, please visit" & Chr$(13) & _
           "http://www.jammingsignal.com " & Chr$(13) & Chr$(13) & _
           "Many thanks to: " & Chr$(13) & _
           "Jeff Ledger, " & _
           "Jeff Hall, " & _
           "Jim Brain, " & _
           "Sysop Fox-1, " & _
           "Eric Pearson, " & _
           "Oliver VieBrooks, " & _
           "John Ward, " & _
           "Rick Lucas, " & _
           "Mike Martin, " & _
           "Moise, " & _
           "CyberJank, " & _
           "Tom Luff, " & _
           "Dustin Chambers, " & _
           "Brian Green, " & _
           "Andrew Wiskow, " & _
           "and the the rest of the gang at " & _
           "forms.petscii.com!", vbInformation
End Sub

Private Sub AdvancedButton_Click()
    Advanced.Show vbModal
End Sub

Private Sub ApplyChanges_Click()
    LogDisplay.Clear
    ApplyINI
    SaveINI
End Sub

Private Sub BoardUp_Click(Index As Integer)
    If Index = 0 Then
        ShowWaitingState
    Else
        ShowOfflineState
    End If
End Sub

Private Sub ClearOnly_Click()
    LogDisplay.Clear
End Sub

Private Sub Exit_Click()

    'Confirm that the Sysop really wants to shut the server down
    Dim Result As Integer
    
    Result = MsgBox("Are you sure you want to shut down the server?", vbYesNo Or vbExclamation)
    
    If (Result = vbYes) Then 'Yes
        ProgramShutDown = True
        DeleteIcon Me    'Get rid of the tray icon
        DoEvents
        Disconnect       'Hang up and disconnect Telnet caller
        End              'Exit the program.
    End If
    
    If (Result = vbNo) Then 'No
        Exit Sub   'Do nothing
    End If
End Sub

Private Sub ForceDisconnect_Click()
    AddMessage "Sysop forced disconnect."
    Disconnect
End Sub

Private Sub Form_Load()

On Error GoTo LoadError:

    ProgramShutDown = False

    'Leif's debug paths  (set by command line within VB IDE, under Project, Properties Make)
    If Command = "-leifdir" Then
      'ChDir "C:\Documents and Settings\Administrator\Desktop\BBS Server 1.4 Testing"
      ChDir "C:\Documents and Settings\Administrator\My Documents\Commodore\BBS Server\"
    End If

    If Command = "-rts" Then
       ForceRTS = True
    Else
       ForceRTS = False
    End If

    'Avoid duplicate instances (from same directory - other instances OK)
    If (App.PrevInstance) Then End
    
    'Load configuration
    LoadINI
    
    'Determine all IP Addresses
    DetermineIPs
    
    'Apply Settings
    ApplyINI
    
    'Show that server is waiting for a call
    ShowWaitingState
    
    'Set that we're not waiting for a special character
    WaitingForResponse = False
    
    'Create Tray Icon
    CreateIcon Me, Me.Caption
    
    'Reset the virtual modem
    ResetVirtualModem
        
    Exit Sub
    
LoadError:
    AddMessage "Form_Load(): " & Err.Description & " (" & Err.Number & ")"
    Resume Next
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    'Sysop Pressed close button - hide instead.
    If (UnloadMode = vbFormControlMenu) Then
        Me.Hide
        Cancel = 1    'Prevent Exit
        Exit Sub
    End If
    
    'All other cases
    ProgramShutDown = True
    DeleteIcon Me    'Get rid of the tray icon
    DoEvents
    Disconnect       'Hang up and disconnect Telnet caller
    End              'Exit the program.
End Sub

Private Sub IPAddressList_Click()
    IPAddressToUse = IPAddressList.text
End Sub


'Handle data from the BBS, received from the Serial port.
Private Sub MSComm_OnComm()
   
    'Update the serial gfx
    UpdateLEDs

On Error GoTo CommError:

    Dim CEvent As Integer
    CEvent = MSComm.CommEvent
    
    Dim Received As String  'Data received over serial port

    Select Case CEvent
        ' Handle each event or error by placing
        ' code below each case statement
        
        ' Errors
        Case comEventBreak      ' A Break was received.
            If (Advanced.DetailedDiagnostics.value) Then
                AddMessage "RS232 Break Received"
            End If
            
            If (Advanced.HangupOnBreak.value) Then
                Disconnect
            End If
            
            Exit Sub
            
        Case comEventFrame      ' Framing Error
            If (Advanced.DetailedDiagnostics.value) Then AddMessage "RS232 Error: Framing Error! (Check Baud Rate)"
            Exit Sub
            
        Case comEventOverrun    ' Data Lost.
            AddMessage "RS232 Error: Overrun!  Data Lost."
            Exit Sub
            
        Case comEventRxOver     ' Receive buffer overflow.
            AddMessage "RS232 Error: Receive Buffer Overflow!"
            Exit Sub
            
        Case comEventRxParity   ' Parity Error.
            AddMessage "RS232 Error: Parity Error!"
            Exit Sub
            
        Case comEventTxFull     ' Transmit buffer full.
            AddMessage "RS232 Error: Transmit Buffer Full!"
            AddMessage "Output buffer size is " & MSComm.OutBufferSize
            Exit Sub
            
        Case comEventDCB        ' Unexpected error retrieving DCB
            AddMessage "RS232 Error: EventDCB!"
            Exit Sub
        
        ' Events
        Case comEvCD            ' Change in the CD line.
            If (MSComm.CDHolding = False) And (ConnectionActive) And (Advanced.CheckDCD.value) Then
                AddMessage "BBS closed connection via DCD, closing Telnet session " & Telnet.RemoteHostIP
                Disconnect
            End If
            Exit Sub
            
        Case comEvCTS           ' Change in the CTS line - do nothing
            Exit Sub
        
        Case comEvDSR           ' Change in the DSR line.
            If (MSComm.DSRHolding = False) And (ConnectionActive) And (Advanced.CheckDSR.value) Then
                AddMessage "BBS closed connection via DSR, closing Telnet session " & Telnet.RemoteHostIP
                Disconnect
                Exit Sub
            End If
            Exit Sub
            
        Case comEvRing          ' Change in the Ring Indicator.
            If (Advanced.DetailedDiagnostics.value) Then AddMessage "Ring Indicator Changed."
            Exit Sub
                        
'----------------------------------------------------------------
'This is the most important event - process a received character
'----------------------------------------------------------------
        Case comEvReceive       ' Received RThreshold # of chars.
            SerialRX.Tag = 255  ' Full Brightness on LED
            Received = MSComm.Input
                        
            'Forward to telnet port if connected and not in Hayes Command Mode
            If (ConnectionActive) And (Not VirtualModem.CommandMode) And (Telnet.State = sckConnected) Then
                
                'ASCII/PETSCII Translation?
                If (UsingCharTranslation) Then
                    Telnet.SendData TranslateSerial(Received)
                Else
                    Telnet.SendData Received
                End If
                
                TelnetTX.Tag = 255 'Full Brightness
                
                'Check for +++, which is done while online
                HandlePluses Received
                
            ' BBS is offline or in command mode, so this may be a command string.
            Else
                'Cancel the outgoing call if a character was received before connect
                If (VirtualModem.OutGoingCall) And Not (Telnet.State = sckConnected) Then
                    Disconnect
                    OutgoingCallDisconnected
                    AddMessage "Outgoing call cancelled."
                End If
                
                If (Advanced.EnableHayes) Then HandleCommandMode Received
            End If
            
            Exit Sub
'----------------------------------------------------------------

        Case comEvSend          ' There are SThreshold number of
                                ' characters in the transmit
                                ' buffer.
            
        Case comEvEOF           ' An EOF charater was found in
                                ' the input stream
                                
    End Select

    AddMessage "Unhandled CommEvent " & CEvent
    
    Exit Sub
    
CommError:
    AddMessage "MSComm_OnComm(): " & Err.Description & " (" & Err.Number & ")"
    Disconnect
    Exit Sub
End Sub

Private Sub OffHookTimer_Timer()
    ' If selected, the virtual modem will stay off-hook until this timer fires.
    OffHookTimer.Enabled = False

    ShowWaitingState
    VirtualModem.OffHook = False
    CompleteDisconnect
End Sub

Private Sub SecondTimer_Timer()

    'Has there been a one-second delay in the data?  If not, then no pause.  (Used by Hayes Emulation)
    If (VirtualModem.DataReceivedinLastSecond) Then
        VirtualModem.OneSecondPauseOccured = False
    Else
        VirtualModem.OneSecondPauseOccured = True
    End If
    
    'Reset for another second
    VirtualModem.DataReceivedinLastSecond = False
    
    If (VirtualModem.WaitingForATA) Then
        If (VirtualModem.ATATimer <= 10) Then
            VirtualModem.ATATimer = VirtualModem.ATATimer + 1
        Else
            AddMessage "Error: BBS Failed to issue ATA!"
            TelnetBBS.Disconnect
            VirtualModem.WaitingForATA = False
        End If
    Else
        VirtualModem.ATATimer = 0
    End If
    
    
    ' Check up on the caller status
    If (ConnectionActive = False) Then
        CallerStatus.Caption = "..."
        Exit Sub
    End If
    
    ' Auto Disconnect if the usder has been on for way too long (default 3 hours)
    If (ConnectionTime <= Val(Advanced.OnlineDisconnectTime.text) * 60) Then
        ConnectionTime = ConnectionTime + 1
    Else
        If (Advanced.OnlineAutoDisconnect.value) Then
            TelnetSendString
            TelnetSendString "Time limit exceeded!  Thanks for calling - call back soon."
            AddMessage "Time limit exceeded - disconnected caller."
            Disconnect
            Exit Sub
        End If
    End If
    
    'Auto disconnect after caller has been idle for too long (default 15 minutes)
    'This can also happen if the BBS hands (user disconnects but BBS doesn't hang up)
    
    If (IdleTime <= Val(Advanced.IdleDisconnectTime.text) * 60) Then
        IdleTime = IdleTime + 1
    Else
        If (Advanced.IdleAutoDisconnect.value) Then
            TelnetSendString
            TelnetSendString "Idle timeout.  Thanks for calling."
            AddMessage "Idle timeout - disconnected caller."
            Disconnect
            Exit Sub
        End If
    End If

    CallerStatus.Caption = "Online " & MinsSecs(ConnectionTime) & " / Idle " & MinsSecs(IdleTime)
End Sub

'Used in debugging
Private Sub SerialDTR_Click()
    MSComm.DTREnable = Not MSComm.DTREnable
    UpdateLEDs
    AddMessage "Manually changed DTR to " & MSComm.DTREnable
End Sub

'Used in debugging
Private Sub SerialRTS_Click()

    If MSComm.Handshaking = comRTS Then
        AddMessage "Can't set RTS with hardware flow control enabled."
    Else
        MSComm.RTSEnable = Not MSComm.RTSEnable
        UpdateLEDs
        AddMessage "Manually changed RTS to " & MSComm.RTSEnable
    End If
End Sub

Private Sub SetAwayMessage_Click()
    AwayMessage.Show vbModal
End Sub

Private Sub Telnet_Close()
    AddMessage "Telnet session with " & Telnet.RemoteHostIP & " disconnected."
    Disconnect
    
    If (Advanced.SendNoCarrier.value) Or (VirtualModem.OutGoingCall) Then
        OutgoingCallDisconnected
    End If
End Sub

Private Sub Incoming_ConnectionRequest(Index As Integer, ByVal requestID As Long)

On Error GoTo RequestError:

    'Invite caller back if a connection is already active
    If (ConnectionActive) Or (VirtualModem.WaitingForATA) Or (VirtualModem.OffHook = True) Then
    
        'Accept this extra connection on a new Winsock instance
        Load Incoming(Incoming.UBound + 1)
        Incoming(Incoming.UBound).Accept requestID
    
        'Record in message log
        AddMessage "Connection from " & Incoming(Incoming.UBound).RemoteHostIP & " but already busy."
        
        'Invite caller back and show current user's online/idle time
        Incoming(Incoming.UBound).SendData BBSNameText.text & Chr$(13) & Chr$(10)
        Incoming(Incoming.UBound).SendData Chr$(13) & Chr$(10)
        Incoming(Incoming.UBound).SendData "Sorry, the BBS is busy.  Please try again in a few minutes!" & Chr$(13) & Chr$(10)
        Incoming(Incoming.UBound).SendData Chr$(13) & Chr$(10)
        Incoming(Incoming.UBound).SendData "Current Caller Status: " & CallerStatus.Caption & Chr$(13) & Chr$(10)
        DoEvents      ' This statement is absolutely required!!!!
        SendBusyText Incoming.UBound  ' Send the contents of busy.txt to caller
        Incoming(Incoming.UBound).Close
        
        'Unload the new instance
        Unload Incoming(Incoming.UBound)
        Exit Sub
    End If
    
    'Invite caller back if board is down
    
    If (BoardUp(1) = True) Then
        'Accept this connection on a new Winsock instance
        Load Incoming(Incoming.UBound + 1)
        Incoming(Incoming.UBound).Accept requestID
    
        'Record in message log
        AddMessage "Connection from " & Incoming(Incoming.UBound).RemoteHostIP & " but BBS is down."
        
        Incoming(Incoming.UBound).SendData Chr$(13) & Chr$(10)
        Incoming(Incoming.UBound).SendData AwayMessage.Message
        Incoming(Incoming.UBound).SendData Chr$(13) & Chr$(10)
        Incoming(Incoming.UBound).SendData Chr$(13) & Chr$(10)
        DoEvents   ' This statement is absolutely required!!!!
        SendBusyText Incoming.UBound
        Incoming(Incoming.UBound).Close
        
        'Unload the new instance
        Unload Incoming(Incoming.UBound)
        Exit Sub
    End If
    
    'Make sure port is closed before accepting
    If (Telnet.State <> sckClosed) Then Telnet.Close
    
    'Reset Timers
    ConnectionTime = 0
    IdleTime = 0
    
    'Accept the request (completes connection)
    Telnet.Accept requestID
    
    'Send the BBS Name
    TelnetSendString BBSNameText.text & vbCrLf
    DoEvents
    
    'Discard any extraneous characters
    Dim temp As String
    Telnet.GetData temp, vbString
    
    'Update status indicators
    ShowOnlineState "Caller from " & Telnet.RemoteHostIP
    
    'Inhibit Sysop from taking board down while caller is connected
    BoardUp(1).Enabled = False
    
    AddMessage "Connection from " & Telnet.RemoteHostIP
    
    If (Advanced.UseCharTranslation) Then 'Prompt for Terminal Selection
         TelnetSendString
         TelnetSendString "please select terminal type:"
         TelnetSendString
         TelnetSendString "1) ascii (standard telnet, default)"
         TelnetSendString "2) petscii c/g (cgterm, cbmterm, c64)"
         
         WaitingForResponse = True
        
         'The program now waits for the Telnet caller to send a response character, which will
         'trigger a call to CompleteConnection(), below.
         Exit Sub
    Else
        CompleteConnection "0"
        Exit Sub
    End If
    
    Exit Sub
    
RequestError:
    AddMessage "Incoming_ConnectionRequest(): " & Err.Description & " (" & Err.Number & ")"
    Resume Next
End Sub

Private Sub CompleteConnection(Terminal As String)

On Error GoTo CompleteError:

    ' Use cases, for future translation file types, or other options.
    
    Select Case Terminal
    
        Case "0":   'No translation options
            UsingCharTranslation = False
            TelnetSendString "connecting..."
    
        Case "1":
            AddMessage "Caller selected ASCII mode."
            UsingCharTranslation = False
            TelnetSendString
            TelnetSendString "Connecting using ASCII mode."
    
        Case "2":
            AddMessage "Caller selected PETSCII mode."
            UsingCharTranslation = True
            SetupTranslation
            TelnetSendString
            TelnetSendString "cONNECTING USING petscii TRANSLATION..."
            
        Case Else:
            TelnetSendString "unknown response, try again"
            WaitingForResponse = True
            
            'Abort connection
            Exit Sub
    End Select
    
    'Call BBS
    ConnectToBBS
       
    'Remain connected until Disconnect() is called.
    Exit Sub
    
CompleteError:
    AddMessage "CompleteConnection(): " & Err.Description & " (" & Err.Number & ")"
    Resume Next
End Sub

Private Sub TelnetSendString(Optional ToSend As String = "", Optional CRLF As Boolean = True, Optional User As Integer = 1)
On Error GoTo SendstringError:

    If (Telnet.State = sckConnected) Then
        Telnet.SendData ToSend
        If (CRLF) Then Telnet.SendData Chr$(13) & Chr$(10)
    End If
    
    Exit Sub
    
SendstringError:
    AddMessage "TelnetSendString(): " & Err.Description & " (" & Err.Number & ")"
    Resume Next
End Sub

Private Sub Telnet_Connect()
    If (VirtualModem.OutGoingCall) Then
        If Advanced.DTROutbound Then MSComm.DTREnable = True
        If Advanced.RTSOutbound Then MSComm.RTSEnable = True
        CallConnected
    End If
End Sub

Private Sub Telnet_DataArrival(ByVal bytestotal As Long)
    'Data has arrived over Telnet
    
On Error GoTo TelnetArrivalError:

    Dim Received As String
    Dim sResponse As String  'For Telnet IAC responses
    TelnetRX.Tag = 255  ' Full Brightness
    
    ' Receive the data into a string.
    Telnet.GetData Received, vbString, bytestotal
    
    
    'If an outgoing Telnet call, may have to sort out IAC characters.
    
    If (VirtualModem.OutGoingCall) And (Not FirstCharReceived) Then
        ' This is the first character.  Check if it's 255.  If so, this is a true Telnet session.
        If Asc(Left$(Received, 1)) = 255 Then
            TrueTelnetMode = True
        End If
        
        FirstCharReceived = True  ' Don't check future characters
    End If
    
    
    If (VirtualModem.OutGoingCall And TrueTelnetMode) Then
    
        'Construct a response if there are any IAC commands
        sResponse = IACResponse(Received)

        If (sResponse <> "") Then
            Telnet.SendData sResponse
            DoEvents
        End If
        
        'Remove all IAC commands
        Received = IACFilter(Received)
    End If 'Outgoing
   
    'Intercept certain response characters at startup
    If (WaitingForResponse) Then
        WaitingForResponse = False
        CompleteConnection (Mid$(Received, 1, 1))
        Exit Sub
    End If
    
    'Echo characters back to Telnet client if required
    If (Advanced.EchoTelnetChars.value) Then Telnet.SendData Received
        
    'Forward to COM port
    If (ConnectionActive) Then
        If (UsingCharTranslation) Then
            If (MSComm.PortOpen) Then MSComm.Output = TranslateTelnet(Received, bytestotal)
        Else
            If (MSComm.PortOpen) Then MSComm.Output = Received
        End If
        
        SerialTX.Tag = 255  ' Full Brightness
    End If

    ' Reset Idle Time
    IdleTime = 0
    Exit Sub
    
TelnetArrivalError:
    AddMessage "Telnet_DataArrival(): " & Err.Description & " (" & Err.Number & ")"
    Disconnect 'Break connection - could be handled better
    Exit Sub
End Sub

Private Sub ConnectToBBS()

On Error GoTo ConnectBBSError:

    'Open the port
    If (MSComm.PortOpen = False) Then
        If (Advanced.DetailedDiagnostics.value) Then AddMessage "Opening COM Port"
        MSComm.PortOpen = True
    End If
    
    'Raise RTS
     If (Advanced.RTSOnConnect.value) Then
        MSComm.RTSEnable = True
        UpdateLEDs
    End If
    
    'Raise DTR
    If (Advanced.DTROnConnect.value) Then
        MSComm.DTREnable = True
        UpdateLEDs
    End If
    
    'Send 'RING'
    If (Advanced.SendRing.value) Then
        MSComm.Output = "RING" & vbCrLf
    Else
        'Send a carriage return to wake BBS up
        MSComm.Output = Chr$(13)
    End If
    
    If (Advanced.PlayWAVonConnect.value) Then
        PlaySoundX Advanced.ConnectionWAV
    End If
    
    If (Advanced.WaitForATA.value) Then
         VirtualModem.WaitingForATA = True
    Else
        ConnectionActive = True
    End If
    
    'Turn off Command Mode
    VirtualModem.CommandMode = False
    
    DoEvents
    
    Exit Sub

ConnectBBSError:
    AddMessage "ConnectToBBS(): " & Err.Description & " (" & Err.Number & ")"
    Resume Next
End Sub

Public Sub SerialReset()

On Error GoTo SerialResetError:
      
    ' Set which COM Port is in use.
    If (MSComm.PortOpen) Then MSComm.PortOpen = False
    
    ' New for 1.2a - override the default of 512 bytes for buffer
    MSComm.OutBufferSize = 10000
    SerialBufCountMax = 0
    
    MSComm.CommPort = Val(COMportText.text)
    If (Advanced.DetailedDiagnostics.value) Then AddMessage "Using COM Port " & MSComm.CommPort
    
    ' Originally, 1200 baud, no parity, 8 data, and 1 stop bit.
    MSComm.Settings = Advanced.SerialSetup.text
    
    ' Tell the control to read entire buffer when Input is used.
    MSComm.InputLen = 0
    
    ' Close the serial port
    If (Advanced.EnableHayes) Then
       If (Advanced.DetailedDiagnostics.value) Then AddMessage "Keeping COM Port Open (Hayes Emulation)"
       If (Not MSComm.PortOpen) Then MSComm.PortOpen = True
    Else
       If (Advanced.DetailedDiagnostics.value) Then AddMessage "Closing COM Port"
       If (MSComm.PortOpen) Then MSComm.PortOpen = False
    End If
    
    'Turn off DTR and RTS
    If (Advanced.LowerRTSonDisconnect) Then
        If (Advanced.DetailedDiagnostics.value) Then AddMessage "Lowering RTS"
        MSComm.RTSEnable = False
    End If
    
    If (Advanced.LowerDTRonDisconnect) Then
        If (Advanced.DetailedDiagnostics.value) Then AddMessage "Lowering DTR"
        MSComm.DTREnable = False
    End If
    
    'Set receive character threshold to 1 (respond to every character)
    MSComm.RThreshold = 1
    
    'Optionally, Re-Enable DTR and RTS after a short delay.  (Fox)
    If (Advanced.ReEnableDTR.value) Or (Advanced.ReEnableDTR.value) Then
        If (Advanced.DetailedDiagnostics.value) Then AddMessage Advanced.CarrierDropDelayTime.text & " seconds delay"
        DoEvents
        Sleep (Advanced.CarrierDropDelayTime.text * 1000)
        DoEvents
        If (MSComm.PortOpen = False) Then MSComm.PortOpen = True
        AddMessage "Reopening COM" & MSComm.CommPort
    End If

    If (Advanced.ReEnableDTR.value) Then
        If (Advanced.DetailedDiagnostics.value) Then AddMessage "Re-enabling DTR"
        If (Advanced.DTROnConnect.value) Then MSComm.DTREnable = True
    End If

     If (Advanced.ReEnableRTS.value) Then
        If (Advanced.DetailedDiagnostics.value) Then AddMessage "Re-enabling RTS"
        If (Advanced.RTSOnConnect.value) Then MSComm.RTSEnable = True
    End If

    'Set Hardware Flow Control
    If (VirtualModem.HardwareFlowControl = True) Then
        MSComm.Handshaking = comRTS
        If (Advanced.DetailedDiagnostics) Then AddMessage "Hardware Flow Control enabled."
    Else
        MSComm.Handshaking = comNone
        If (Advanced.DetailedDiagnostics) Then AddMessage "Hardware Flow Control disabled."
    End If
    
    'New - force RTS
    If (ForceRTS) Then
       MSComm.RTSEnable = True
       AddMessage "RTS forced on (-rts option)."
    End If
    
    UpdateLEDs
    
    Exit Sub
    
SerialResetError:
    AddMessage "SerialReset(): " & Err.Description & " (" & Err.Number & ")"
    Resume Next
End Sub

Private Sub Telnet_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    AddMessage "Winsock Error " & Number & "-" & Description
    
     Disconnect 'Break connection
    
    If (VirtualModem.OutGoingCall) Then
        If (MSComm.PortOpen And Advanced.SendWinsockErrors.value) Then
            SerialSendString Description
            DoEvents
        End If
        
        OutgoingCallDisconnected
    End If
End Sub

Private Sub GraphicsTimer_Timer()
   
    ' A fun bit of color showing the data flow.
    ' The Ifs avoid the annoying VB 'flicker'
    Const MIN = 10
   
On Error GoTo TimerError:

    'Track the buffer size
    Dim SerialBufCountNow As Integer
    SerialBufCountNow = MSComm.OutBufferCount
    If (SerialBufCountNow > SerialBufCountMax) Then SerialBufCountMax = SerialBufCountNow
    OutCount.Caption = "Output Buffer Count: " & SerialBufCountNow & "  Maximum: " & SerialBufCountMax

    With TelnetRX
        If (Val(.Tag) > MIN) Then
            .Tag = Val(.Tag) * 0.9
            .BackColor = RGB(0, Val(.Tag), 0)
        Else
            .Tag = 0
        End If
    End With
    
    With SerialTX
         If (Val(.Tag) > MIN) Then
            .Tag = Val(.Tag) * 0.9
            .BackColor = RGB(0, Val(.Tag), 0)
        Else
            .Tag = 0
        End If
    End With
    
    With SerialRX
        If (Val(.Tag) > MIN) Then
            .Tag = Val(.Tag) * 0.9
            .BackColor = RGB(Val(.Tag), 0, 0)
        Else
            .Tag = 0
        End If
    End With
    
   With TelnetTX
      If (Val(.Tag) > MIN) Then
            .Tag = Val(.Tag) * 0.9
            .BackColor = RGB(Val(.Tag), 0, 0)
        Else
            .Tag = 0
        End If
    End With
    
    Exit Sub
    
TimerError:
    AddMessage "Timer(): " & Err.Description & " (" & Err.Number & ")"
    Resume Next
End Sub

'Disconnects - Closes Telnet Port, Resets Serial Port, Clears ConnectionActive and other flag
Public Sub Disconnect()

On Error GoTo DisconnectError:

    If (MSComm.PortOpen) And (Advanced.SendCtrlC.value) Then
        If (Advanced.DetailedDiagnostics.value) Then AddMessage "Sending CTRL-C to BBS before Disconnect"
        MSComm.Output = Chr$(3)
        DoEvents
    End If
    
    'Set internal connection state to false
    ConnectionActive = False
    
    'Clear flag for waiting for characters
    WaitingForResponse = False
                
    'Reset COM Port
    SerialReset
    
    'Close Telnet session
    Telnet.Close
    
    'Give a Beep
    If (Advanced.PlayWAVonDisconnect.value) And (Not ProgramShutDown) And (Not VirtualModem.OutGoingCall) Then
        PlaySoundX Advanced.DisconnectWAV
    End If
    
    'Optionally, stay off-hook for a few more seconds to give BBS time to reinitialize
    If (Not VirtualModem.OutGoingCall) And (Advanced.StayOffHook.value) Then
        VirtualModem.OffHook = True
        ShowOffHookState
        AddMessage "Staying off hook for " & Val(Advanced.StayOffHookTime.text) & " seconds"
        TelnetBBS.Refresh
        DoEvents
        OffHookTimer.Interval = Val(Advanced.StayOffHookTime.text) * 1000 'milliseconds
        OffHookTimer.Enabled = True
    Else
        CompleteDisconnect
    End If
    
    'Disconnect sequence continues in CompleteDisconnect(), below.
    
    Exit Sub
    
DisconnectError:
    AddMessage "Disconnect(): " & Err.Description & " (" & Err.Number & ")"
    Exit Sub

End Sub

Public Sub CompleteDisconnect()

On Error Resume Next
 
    'Reset Telnet Server for the Next Caller.
    Telnet.LocalPort = 0
    Telnet.Listen
    
    'Show new status.   If BBS was 'down', it stays down.
    If (BoardUp(0).value = True) Then
        ShowWaitingState
    Else
        ShowOfflineState
    End If
    
    'Not in Hayes Command mode.
    VirtualModem.CommandMode = False
    
    'We have not received any characters.
    FirstCharReceived = False
    
    'We don't know what Telnet mode is next, so set to False
    TrueTelnetMode = False
    
    'Allow BBS to be taken down.
    BoardUp(1).Enabled = True
    
    Exit Sub
    
CompleteDisconnectError:
    AddMessage "CompleteDisconnect(): " & Err.Description & " (" & Err.Number & ")"
    Exit Sub
End Sub


Private Function MinsSecs(Seconds As Long) As String
    Dim temp1 As Long
    Dim temp2 As Long

    temp1 = Fix(Seconds / 60)
    temp2 = Seconds - (temp1 * 60)
    
    MinsSecs = temp1 & ":" & Format(temp2, "00")
End Function

Public Sub ShowWaitingState()
    BBSState.Caption = "WAITING"
    BBSState.BackColor = vbYellow
    TelnetBBS.Icon = WaitingIcon.Picture
    RefreshTrayIcon "Waiting for Call"
    ForceDisconnect.Visible = False
End Sub

Private Sub ShowOnlineState(text As String)
    BBSState.Caption = "ONLINE"
    BBSState.BackColor = vbGreen
    TelnetBBS.Icon = OnlineIcon.Picture
    RefreshTrayIcon text
    ForceDisconnect.Visible = True
End Sub

Public Sub ShowOfflineState()
    BBSState.Caption = "OFFLINE"
    BBSState.BackColor = vbRed
    TelnetBBS.Icon = OfflineIcon.Picture
    RefreshTrayIcon "BBS is down!"
    ForceDisconnect.Visible = False
End Sub

Public Sub ShowOutgoingState()
    BBSState.Caption = "OUTGOING"
    BBSState.BackColor = vbCyan
    TelnetBBS.Icon = OnlineIcon.Picture
    RefreshTrayIcon "Outgoing call in progress"
    ForceDisconnect.Visible = False
End Sub

Public Sub ShowOffHookState()
    BBSState.Caption = "OFF HOOK"
    BBSState.BackColor = RGB(255, 100, 100)
    TelnetBBS.Icon = OfflineIcon.Picture
    RefreshTrayIcon "BBS is off the hook"
    ForceDisconnect.Visible = False
End Sub

'Handle events on the System Tray Icon
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Dim Result As Long
    Dim msg As Long

    If Me.ScaleMode = vbPixels Then
        msg = X
    Else
        msg = X / Screen.TwipsPerPixelX
    End If
        
    'This was in the original code, but we don't need it.  Kept for future.
    'If msg = 517 Then
        'Me.PopupMenu mnu_file
    'Else
    
    If (msg = DOUBLECLICK) Then   'Double-clicked on Icon
       Me.Show
    End If
End Sub

Private Sub RefreshTrayIcon(text As String)
    If (Not ProgramShutDown) Then
        DeleteIcon Me
        CreateIcon Me, Me.Caption & " - " & text
    End If
End Sub

'Update the serial activity LEDs
Private Sub UpdateLEDs()
    'Serial Inputs
    SerialDCD.BackColor = Abs(vbRed * MSComm.CDHolding)
    SerialDSR.BackColor = Abs(vbRed * MSComm.DSRHolding)
    SerialCTS.BackColor = Abs(vbRed * MSComm.CTSHolding)
    
    'Serial Outputs
    SerialDTR.BackColor = Abs(vbGreen * MSComm.DTREnable)
    
    If (MSComm.Handshaking = comRTS) Then
        SerialRTS.BackColor = RGB(0, 200, 0)
    Else
        SerialRTS.BackColor = Abs(vbGreen * MSComm.RTSEnable)
    End If
End Sub

Private Sub SendBusyText(Index As Integer)
    
On Error GoTo NoFile:
    
    Dim temp As String
    
    Close #5
    Open "busy.txt" For Input As #5
    
    While Not EOF(5)
        Line Input #5, temp
        Incoming(Index).SendData temp & Chr$(13) & Chr$(10)
        DoEvents
    Wend
    
    Close #5
    Exit Sub

NoFile:
    Exit Sub ' No need for error msg
End Sub

