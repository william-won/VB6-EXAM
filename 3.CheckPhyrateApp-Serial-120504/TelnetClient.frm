VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{9F3B4DE1-AA29-11D1-A3D9-FDA4E35D1D25}#1.0#0"; "Io.ocx"
Begin VB.Form frmTelnet 
   BackColor       =   &H00C0FFFF&
   Caption         =   "CheckPhyRateApp v1.0     Copyright (c) ATAW"
   ClientHeight    =   6030
   ClientLeft      =   2535
   ClientTop       =   2445
   ClientWidth     =   6360
   FillColor       =   &H00800000&
   FillStyle       =   2  '수평선
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "TelnetClient.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6030
   ScaleWidth      =   6360
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Height          =   3615
      Left            =   120
      TabIndex        =   20
      Top             =   1200
      Width           =   4095
      Begin VB.TextBox txtRxPhyRate 
         Alignment       =   2  '가운데 맞춤
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   2160
         TabIndex        =   25
         Text            =   "-"
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox txtTxPhyRate 
         Alignment       =   2  '가운데 맞춤
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         TabIndex        =   24
         Text            =   "-"
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox txtResult 
         Alignment       =   2  '가운데 맞춤
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   23
         Text            =   "Result"
         Top             =   2520
         Width           =   3855
      End
      Begin VB.TextBox txtMAC 
         Alignment       =   2  '가운데 맞춤
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "-"
         Top             =   480
         Width           =   2535
      End
      Begin VB.TextBox txtMode 
         Alignment       =   2  '가운데 맞춤
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "-"
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "RX PhyRate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   29
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Tx PhyRate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0FFFF&
         Caption         =   "MAC Address"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Mode"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   26
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.OptionButton OptionSerial 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Serial"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   19
      Top             =   600
      Value           =   -1  'True
      Width           =   1575
   End
   Begin VB.OptionButton OptionTCPIP 
      BackColor       =   &H00C0FFFF&
      Caption         =   "TCP IP"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10800
      TabIndex        =   18
      Top             =   1020
      Width           =   1695
   End
   Begin VB.ComboBox ComboPort 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   240
      TabIndex        =   17
      Text            =   "COM1"
      Top             =   600
      Width           =   1815
   End
   Begin VB.ComboBox ComboBaud 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   2280
      TabIndex        =   16
      Text            =   "4800"
      Top             =   600
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "수동명령"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   4440
      TabIndex        =   10
      Top             =   1200
      Width           =   1815
      Begin VB.CommandButton Disconnect 
         BackColor       =   &H00C0FFFF&
         Caption         =   "연결끊기"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         MaskColor       =   &H00C0FFFF&
         Style           =   1  '그래픽
         TabIndex        =   13
         Top             =   1560
         Width           =   1575
      End
      Begin VB.CommandButton cmdinfo 
         BackColor       =   &H00C0FFFF&
         Caption         =   "속도 조회"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         Style           =   1  '그래픽
         TabIndex        =   12
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton Connect 
         BackColor       =   &H00C0FFFF&
         Caption         =   "연결하기"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         Style           =   1  '그래픽
         TabIndex        =   11
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.TextBox txtRawRate 
      Alignment       =   2  '가운데 맞춤
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   4920
      TabIndex        =   8
      Text            =   "5"
      Top             =   3960
      Width           =   855
   End
   Begin VB.CheckBox CheckAutoConnect 
      BackColor       =   &H00C0FFFF&
      Caption         =   "자동 연결및 조회"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6540
      TabIndex        =   7
      Top             =   360
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4680
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtPort 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   9240
      TabIndex        =   4
      Text            =   "40000"
      Top             =   900
      Width           =   1215
   End
   Begin VB.TextBox txtRemoteAddress 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   6600
      TabIndex        =   3
      Text            =   "192.168.0.106"
      Top             =   900
      Width           =   2535
   End
   Begin VB.TextBox txtTime 
      Alignment       =   2  '가운데 맞춤
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2640
      TabIndex        =   2
      Text            =   "09:00:00"
      Top             =   5010
      Width           =   1455
   End
   Begin VB.TextBox txtDate 
      Alignment       =   2  '가운데 맞춤
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1200
      TabIndex        =   1
      Text            =   "09/06/29"
      Top             =   5010
      Width           =   1455
   End
   Begin VB.Timer TimerState 
      Interval        =   3000
      Left            =   120
      Top             =   5880
   End
   Begin VB.Timer Timer1sec 
      Interval        =   1000
      Left            =   600
      Top             =   5880
   End
   Begin MSWinsockLib.Winsock WinsockClient 
      Left            =   4320
      Top             =   6000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin ComctlLib.StatusBar stbStatusBar 
      Align           =   2  '아래 맞춤
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   5535
      Width           =   6360
      _ExtentX        =   11218
      _ExtentY        =   873
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   2822
            MinWidth        =   2822
            Text            =   "해제"
            TextSave        =   "해제"
            Key             =   "Mode"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Operating Mode"
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   "None"
            TextSave        =   "None"
            Key             =   "Lip"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Local IP"
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   "None"
            TextSave        =   "None"
            Key             =   "Rip"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Remote IP"
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   2699
            TextSave        =   ""
            Key             =   "Status"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Last Status Message"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin IOLib.IO IO1 
      Left            =   5520
      Top             =   5760
      _Version        =   65536
      _ExtentX        =   1270
      _ExtentY        =   1270
      _StockProps     =   0
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Port"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Baudrate"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   14
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Low Rate Alarm"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   9
      Top             =   3480
      Width           =   2055
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   10080
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Remote IP Port"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   9240
      TabIndex        =   6
      Top             =   540
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Remote IP Address(Host Name)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   6480
      TabIndex        =   5
      Top             =   540
      Width           =   2895
   End
End
Attribute VB_Name = "frmTelnet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Const GO_NORM = 0

Const GO_ESC1 = 1
Const GO_ESC2 = 2
Const GO_ESC3 = 3
Const GO_ESC4 = 4
Const GO_ESC5 = 5

Const GO_IAC1 = 6
Const GO_IAC2 = 7
Const GO_IAC3 = 8
Const GO_IAC4 = 9
Const GO_IAC5 = 10
Const GO_IAC6 = 11


Const SUSP = 237
Const ABORT = 238      'Abort
Const SE = 240         'End of Subnegotiation
Const NOP = 241
Const DM = 242         'Data Mark
Const BREAK = 243      'BREAK
Const IP = 244         'Interrupt Process
Const AO = 245         'Abort Output
Const AYT = 246        'Are you there
Const EC = 247         'Erase character
Const EL = 248         'Erase Line
Const GOAHEAD = 249    'Go Ahead
Const SB = 250         'What follows is subnegotiation
Const WILLTEL = 251
Const WONTTEL = 252
Const DOTEL = 253
Const DONTTEL = 254
Const IAC = 255

Const BINARY = 0
Const ECHO = 1
Const RECONNECT = 2
Const SGA = 3
Const AMSN = 4
Const STATUS = 5
Const TIMING = 6
Const RCTAN = 7
Const OLW = 8
Const OPS = 9
Const OCRD = 10
Const OHTS = 11
Const OHTD = 12
Const OFFD = 13
Const OVTS = 14
Const OVTD = 15
Const OLFD = 16
Const XASCII = 17
Const LOGOUT = 18
Const BYTEM = 19
Const DET = 20
Const SUPDUP = 21
Const SUPDUPOUT = 22
Const SENDLOC = 23
Const TERMTYPE = 24
Const EOR = 25
Const TACACSUID = 26
Const OUTPUTMARK = 27
Const TERMLOCNUM = 28
Const REGIME3270 = 29
Const X3PAD = 30
Const NAWS = 31
Const TERMSPEED = 32
Const TFLOWCNTRL = 33
Const LINEMODE = 34
Const DISPLOC = 35
Const ENVIRON = 36
Const AUTHENTICATION = 37
Const UNKNOWN39 = 39
Const EXTENDED_OPTIONS_LIST = 255
Const RANDOM_LOSE = 256

'------------------------------------------------------------
Private Operating       As Boolean
Private Connected       As Boolean
Public Receiving        As Boolean

Private parsedata(10)   As Integer
Private ppno            As Integer


Private control_on      As Boolean


Public RemoteIPAd  As String
Public RemotePort  As Long

Public TraceTelnet As Boolean
Public Tracevt100   As Boolean

Private sw_ugoahead As Boolean
Private sw_igoahead As Boolean
Private sw_echo     As Boolean
Private sw_linemode As Boolean
Private sw_termsent As Boolean
Private substate    As Boolean

'kate send complete
Private Sendcomplete As Boolean
Private AmrWrSend As Boolean
Private AmrCmdCode As Integer
Private PasswdTryCount As Integer
Private ModeDS2 As Integer

Private CheckAdmin As Boolean
Private WaitAnswer As Boolean
Private SendNext As Integer



Private Sub cmdinfo_Click()
  If Connected Then
    Dim CH As String
    CH = "/i" & vbCrLf
    If (OptionSerial.Value) Then
        Dim Result As Byte
        Dim i As Integer

        Result = IO1.WriteString(CH)
        While IO1.NumCharsOutQue <> 0
        Wend

        If Result Then
            Dim RxMsg As String
            Dim bytesTotal As Long
                
            If IO1.NumCharsInQue = 0 Then IO1.Sleep (100)
            If IO1.NumCharsInQue = 0 Then IO1.Sleep (100)

            bytesTotal = IO1.NumCharsInQue
            RxMsg = IO1.ReadString(IO1.NumCharsInQue)

            RecieveData_Parse RxMsg, bytesTotal
        End If
    Else
        Debug.Print "Tx:" & CH
        WinsockClient.SendData CH
    End If
  Else
    NotConnectedAlarm
  End If
End Sub

Private Sub Connect_Click()
   On Error Resume Next                                   ' Handle errors...
'------------------------------------------------------------
    If (OptionSerial.Value) Then
        Dim Result As Byte
        TimerState.Enabled = True
        IO1.Close
        Result = IO1.Open("COM1:", "baud=4800 parity=N data=8 stop=1")
        IO1.SetTimeOut (50)
        IO1.Mode = 0 '2: Async Mode
        Connected = Result
        Operating = False
    ElseIf Not Operating Then
        Operating = True
      If TraceTelnet Then Debug.Print Int(Timer) & " - [DoConnect] : " & vbCrLf
        With WinsockClient
            If .State <> 0 Then
                .Close
                .RemotePort = 0
                .LocalPort = 0
                Do
                Loop Until .State = 0
            End If
            frmTelnet.WinsockClient.Close
            frmTelnet.WinsockClient.LocalPort = 0
            frmTelnet.RemoteIPAd = txtRemoteAddress
            frmTelnet.RemotePort = txtPort
            frmTelnet.WinsockClient.RemotePort = txtPort
            frmTelnet.WinsockClient.RemoteHost = txtRemoteAddress
            .RemoteHost = RemoteIPAd 'txtRemoteAddress
            .RemotePort = RemotePort 'txtPort
            .Connect ' Attempt new connection
            'term_init
            frmTelnet.stbStatusBar.Panels(4).Text = "연결중"
        End With
    End If

End Sub

Private Sub Disconnect_Click()
    If (OptionSerial.Value) Then
        IO1.Close
    Else
        WinsockClient_Close
    End If
    TimerState.Enabled = False
    txtMAC.Text = "-"
    txtMode.Text = "-"
    txtTxPhyRate.Text = "-"
    txtRxPhyRate.Text = "-"
    txtResult.Text = "Result"
    txtResult.BackColor = &H80000005
End Sub

Private Function NotConnectedAlarm()
  MsgBox "연결되어있지않음", , "알림"
End Function

Private Sub Form_Load()
    Dim i As Integer
    
    'STATSBAR.Caption = ""

    TimerState.Enabled = False
    
    RemoteIPAd = "192.168.0.105"
    RemotePort = 40000
    stbStatusBar.Panels(3).Text = WinsockClient.LocalIP
    'term_init
    txtDate = Format(Date, "yy/mm/dd")
    txtTime = Format(Time, "hh:mm:ss")
    
    ModeDS2 = False
    WaitAnswer = False
    CheckAdmin = True

End Sub

Private Sub Form_Paint()
 'term_redrawscreen
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    With WinsockClient
        .Close                            ' Clear any errors...
        .RemoteHost = "0.0.0.0"
        .RemotePort = 0
    End With
    Operating = False
    Connected = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End                                  ' End program forcefully
End Sub

Private Sub mClose_Click()
    WinsockClient_Close
End Sub

Private Sub mExit_Click()
    End
End Sub

Private Sub OptionTCPIP_Click()
    OptionSerial.Value = False
    OptionTCPIP.Value = True
End Sub

Private Sub OptSerial_Click()
    OptionSerial.Value = True
    OptionTCPIP.Value = False
End Sub

Private Sub Timer1sec_Timer()
    txtDate = Format(Date, "yy/mm/dd")
    txtTime = Format(Time, "hh:mm:ss")
End Sub

Private Sub TimerState_Timer()
    'If CheckAutoConnect Then
        If (Connected = False) And (frmTelnet.stbStatusBar.Panels(4).Text <> "연결중") Then
            Call Connect_Click
        ElseIf Connected Then
            Call cmdinfo_Click
        End If
    'End If
End Sub

Private Sub WinsockClient_sendcomplete()
    Sendcomplete = True
    'If AmrWrSend = True Then
        'AmrWrSend = False
        'cmdAmrRd_Click
    'End If
End Sub

Private Sub WinsockClient_Close()
        
        frmTelnet.stbStatusBar.Panels(1).Text = "해제"
        frmTelnet.stbStatusBar.Panels(3).Text = WinsockClient.LocalIP
        frmTelnet.stbStatusBar.Panels(2).Text = ""
        frmTelnet.stbStatusBar.Panels(4).Text = "연결해제됨"
        
      If TraceTelnet Then Debug.Print Int(Timer) & " - [Closed  ] : Connection Reset By Peer "
        With WinsockClient
            .Close                                     ' Clear any errors...
            .RemotePort = 0
            .LocalPort = 0
        End With
        Operating = False
        Connected = False
 
End Sub

Private Sub WinsockClient_Connect()

Dim ConnectString As String

'------------------------------------------------------------
        
      If TraceTelnet Then Debug.Print Int(Timer) & " - [Connect] : " & _
                    "[" & WinsockClient.RemoteHost & "] " & _
                    "[" & WinsockClient.RemoteHostIP & "] " & _
                    "[" & CStr(WinsockClient.RemotePort) & "]"  ' Display connection info
        
         
        sw_ugoahead = True
        sw_igoahead = False
        sw_echo = True
        sw_linemode = False
        sw_termsent = False
        substate = False
         
        'ConnectString = Chr$(IAC) & Chr$(DOTEL) & Chr$(ECHO) _
        '               & Chr$(IAC) & Chr$(DOTEL) & Chr$(SGA) _
        '               & Chr$(IAC) & Chr$(WILLTEL) & Chr$(NAWS) _
        '               & Chr$(IAC) & Chr$(WILLTEL) & Chr$(TERMTYPE) _
        '               & Chr$(IAC) & Chr$(WILLTEL) & Chr$(TERMSPEED)

        
        'WinsockClient.SendData ConnectString
        
        'If TraceTelnet Then Debug.Print "SENT: DOTEL  ECHO SGA"
        'If TraceTelnet Then Debug.Print "SENT: WILL NAWS TERMTYPE TERMSPEED"
        
        Connected = True
        
        If CheckAdmin Then
          Dim CH As String

          CH = "mode admin" & vbCrLf
          Debug.Print "Tx:" & CH
          WinsockClient.SendData CH
        End If

        txtMAC.Text = "-"
        txtMode.Text = "-"
        txtTxPhyRate.Text = "-"
        txtRxPhyRate.Text = "-"
        txtResult.Text = "Result"
        txtResult.BackColor = &H80000005
        frmTelnet.stbStatusBar.Panels(1).Text = "연결"
        frmTelnet.stbStatusBar.Panels(3).Text = WinsockClient.LocalIP
        frmTelnet.stbStatusBar.Panels(2).Text = WinsockClient.RemoteHostIP
        frmTelnet.stbStatusBar.Panels(4).Text = "연결되었음"
        
End Sub

Private Function xtoi(c As Byte) As Integer
  If ((c >= &H30) And (c <= &H39)) Then
    xtoi = c - &H30
  ElseIf ((c >= &H41) And (c <= &H46)) Then
    xtoi = c - &H37
  ElseIf ((c >= &H61) And (c <= &H66)) Then
    xtoi = c - &H57
  Else
    xtoi = -1
  End If
End Function

Private Sub RecieveData_Parse(rxStr As String, bytesTotal As Long)
  Dim disStr As String
  Dim cChar As String
  Dim i As Integer
  Dim j As Integer
  Dim CH As String
  Dim MACAddress As String
  
  If bytesTotal > 20 Then
    txtDate = Format(Date, "yy/mm/dd") 'txtDate = Format(Date, "yyyy\/mm\/dd")
    txtTime = Format(Time, "hh:mm:ss")
  End If

  If bytesTotal > 9 Then   '무작위 scroll 방지
    i = 1
    Do While i <= bytesTotal
      cChar = Mid(rxStr, i, 1)
      If cChar = vbLf Or cChar = vbCr Then
        disStr = disStr & vbCrLf
        i = i + 1
      Else
        If i < bytesTotal Then
          If (AscB(cChar) = 255) And (StrComp(Mid(rxStr, i + 1, 1), "?") = 0) Then
              i = i + 1 'skip 0xff 0x3f("?")
          Else
            disStr = disStr & cChar
          End If
        Else
          disStr = disStr & cChar
        End If
        i = i + 1
      End If
    Loop

    Debug.Print "[Rx:" & bytesTotal & "bytes]" & disStr
  End If
  
  If InStr(rxStr, "MAC: ") > 0 And InStr(rxStr, "MODE: ") > 0 Then
    j = InStr(rxStr, "MODE: ") + 5
    txtMode = RTrim(Mid(rxStr, j, 3))

    j = InStr(rxStr, "MAC: ") + 5
    txtMAC = RTrim(Mid(rxStr, j, 17))
  End If

  If InStr(rxStr, "Forwarding (M)") > 0 Or InStr(rxStr, "Forwarding (S)") > 0 Then 'Spirit
    '11. 00:0B:29:00:0A:B2     36 Mbps        33 Mbps       Forwarding (M)
    ' 9. 00:07:7F:88:01:57     125 Mbps        10 Mbps       Forwarding (M)
    If (InStr(rxStr, "Forwarding (M)") > 0) Then
        j = InStr(rxStr, "Forwarding (M)")
    Else
        j = InStr(rxStr, "Forwarding (S)")
    End If
    
    
    rxStr = Trim(Mid(rxStr, j - 30, 30))
    j = InStr(rxStr, "Mbps")
    
    txtTxPhyRate.Text = Trim(Mid(rxStr, 1, j - 1))
    rxStr = Trim(Mid(rxStr, j + 4, Len(rxStr) - 4))
    j = InStr(rxStr, "Mbps")
    txtRxPhyRate.Text = Trim(Mid(rxStr, 1, j - 1))

    If ((Val(txtTxPhyRate.Text) <= Val(txtRawRate.Text)) Or (Val(txtRxPhyRate.Text) <= Val(txtRawRate.Text))) Then
        txtResult = "Fail"
        txtResult.BackColor = &HFF&
    Else
        txtResult = "OK"
        txtResult.BackColor = &HFF00&
    End If
  End If
End Sub
Private Sub WinsockClient_DataArrival(ByVal bytesTotal As Long)
  Dim rxStr As String

  WinsockClient.GetData rxStr
  
  RecieveData_Parse rxStr, bytesTotal

End Sub


Private Function iac1(CH As Byte) As Integer

  ' Debug.Print "IAC : ";
  iac1 = GO_NORM

  Select Case CH
    Case DOTEL
      iac1 = GO_IAC2
    Case DONTTEL
      iac1 = GO_IAC6
    Case WILLTEL
      iac1 = GO_IAC3
    Case WONTTEL
      iac1 = GO_IAC4
    Case SB
      iac1 = GO_IAC5
      ppno = 0
    Case SE
      ' End of negotiation string, string is in parsedata()
      Select Case parsedata(0)
        Case TERMTYPE
          If parsedata(1) = 1 Then
               If TraceTelnet Then Debug.Print "SENT: SB TERMTYPE VT100"
                WinsockClient.SendData Chr$(IAC) & Chr$(SB) & Chr$(TERMTYPE) & "DEC-VT100" & Chr$(0) & Chr$(IAC) & Chr$(SE)
          End If
        Case TERMSPEED
          If parsedata(1) = 1 Then
                ' Debug.Print "TERMSPEED"
                If TraceTelnet Then Debug.Print "SENT: SB TERMSPEED 38400"
                WinsockClient.SendData Chr$(IAC) & Chr$(WILLTEL) & Chr$(CH)
                WinsockClient.SendData Chr$(IAC) & Chr$(SB) _
                                & Chr$(TERMSPEED) & Chr$(0) _
                                & "57600,57600" _
                                & Chr$(IAC) & Chr$(SE)
          End If
      End Select
  End Select

End Function

Private Function iac2(CH As Byte) As Integer

  'DO Processing Respond with WILL or WONT

  If TraceTelnet Then Debug.Print "                                                                   RECEIVED DO : ";
  iac2 = GO_NORM

  Select Case CH
    Case BINARY
        If TraceTelnet Then Debug.Print "BINARY"
        WinsockClient.SendData Chr$(IAC) & Chr$(WONTTEL) & Chr$(BINARY)
        If TraceTelnet Then Debug.Print "SENT: WONT BINARY"
    Case ECHO
        If TraceTelnet Then Debug.Print "ECHO"
        WinsockClient.SendData Chr$(IAC) & Chr$(WONTTEL) & Chr$(ECHO)
        If TraceTelnet Then Debug.Print "SENT: WONT ECHO"
    Case NAWS
        If TraceTelnet Then Debug.Print "WINDOW SIZE"
        WinsockClient.SendData Chr$(IAC) & Chr$(SB) & Chr$(NAWS) & Chr$(0) & Chr$(80) & Chr$(0) & Chr$(24) & Chr$(IAC) & Chr$(SE)
        If TraceTelnet Then Debug.Print "SENT: SB WINDOW SIZE 80x24"
    Case SGA
        If TraceTelnet Then Debug.Print "SGA"
        If Not sw_igoahead Then
            If TraceTelnet Then Debug.Print "SENT: WILL SGA"
            WinsockClient.SendData Chr$(IAC) & Chr$(WILLTEL) & Chr$(SGA)
            sw_igoahead = True
        Else
           If TraceTelnet Then Debug.Print "DID NOT RESPOND"
        End If
    Case TERMTYPE
        If TraceTelnet Then Debug.Print "TERMTYPE"
        If Not sw_termsent Then
            If TraceTelnet Then Debug.Print "SENT: WILL TERMTYPE"
              sw_termsent = True
              WinsockClient.SendData Chr$(IAC) & Chr$(WILLTEL) & Chr$(TERMTYPE)
            If TraceTelnet Then Debug.Print "SENT: SB TERMTYPE VT100"
              WinsockClient.SendData Chr$(IAC) & Chr$(SB) & Chr$(TERMTYPE) & _
              Chr$(0) & "VT100" & Chr$(IAC) & Chr$(SE)
         Else
            If TraceTelnet Then Debug.Print "DID NOT RESPOND"
         End If
 
    Case TERMSPEED
        If TraceTelnet Then Debug.Print "TERMSPEED"
        If TraceTelnet Then Debug.Print "SENT: WILL TERMSPEED"
        WinsockClient.SendData Chr$(IAC) & Chr$(WILLTEL) & Chr$(TERMSPEED)
      
    If TraceTelnet Then Debug.Print "SENT: SB TERMSPEED 57600"
        WinsockClient.SendData Chr$(IAC) & Chr$(SB) & Chr$(TERMSPEED) & Chr$(0)
        WinsockClient.SendData "57600,57600"
        WinsockClient.SendData Chr$(IAC) & Chr$(SE)
      
    Case TFLOWCNTRL
        If TraceTelnet Then Debug.Print "TFLOWCNTRL"
        If TraceTelnet Then Debug.Print "SENT: WONT FLOWCONTROL"
        WinsockClient.SendData Chr$(IAC) & Chr$(WONTTEL) & Chr$(CH)
      
    Case LINEMODE
        If TraceTelnet Then Debug.Print "LINEMODE"
        If TraceTelnet Then Debug.Print "SENT: WONT LINEMODE"
        WinsockClient.SendData Chr$(IAC) & Chr$(WONTTEL) & Chr$(CH)
      
    Case STATUS
        If TraceTelnet Then Debug.Print "STATUS"
        If TraceTelnet Then Debug.Print "SENT: WONT STATUS"
        WinsockClient.SendData Chr$(IAC) & Chr$(WONTTEL) & Chr$(CH)
      
    Case TIMING
        If TraceTelnet Then Debug.Print "TIMING"
        If TraceTelnet Then Debug.Print "SENT: WONT TIMING"
        WinsockClient.SendData Chr$(IAC) & Chr$(WONTTEL) & Chr$(CH)
      
    Case DISPLOC
        If TraceTelnet Then Debug.Print "DISPLOC"
        If TraceTelnet Then Debug.Print "SENT: WONT DISPLOC"
        WinsockClient.SendData Chr$(IAC) & Chr$(WONTTEL) & Chr$(CH)
    
    Case ENVIRON
        If TraceTelnet Then Debug.Print "ENVIRON"
        If TraceTelnet Then Debug.Print "SENT: WONT ENVIRON"
        WinsockClient.SendData Chr$(IAC) & Chr$(WONTTEL) & Chr$(CH)
    
    Case UNKNOWN39
        If TraceTelnet Then Debug.Print "UNKNOWN39"
        If TraceTelnet Then Debug.Print "SENT: WONT " & Asc(CH)
        WinsockClient.SendData Chr$(IAC) & Chr$(WONTTEL) & Chr$(CH)
    
    Case AUTHENTICATION
        If TraceTelnet Then Debug.Print "AUTHENTICATION"
        If TraceTelnet Then Debug.Print "SENT: WILL "; AUTHENTICATION; ""
        WinsockClient.SendData Chr$(IAC) & Chr$(WILLTEL) & Chr$(CH)
      
        If TraceTelnet Then Debug.Print "SENT: SB AUTHENTICATION"
        WinsockClient.SendData Chr$(IAC) & _
                          Chr$(SB) & _
                          Chr$(AUTHENTICATION) & _
                          Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & _
                          Chr$(IAC) & _
                          Chr$(SE)
    Case Else
        If TraceTelnet Then Debug.Print "UNKNOWN CMD " & Asc(CH)
        If TraceTelnet Then Debug.Print "SENT: WONT UNKNOWN CMD " & CH
        WinsockClient.SendData Chr$(IAC) & Chr$(WONTTEL) & Chr$(CH)
  End Select

End Function

Private Function iac3(CH As Byte) As Integer

  ' WILL Processing - Respond with DO or DONT
  
If TraceTelnet Then Debug.Print "                                                                   RECEIVED WILL : ";

  iac3 = GO_NORM

  Select Case CH
    Case ECHO
    If TraceTelnet Then Debug.Print "ECHO"
      If Not sw_echo Then
        sw_echo = True
        WinsockClient.SendData Chr$(IAC) & Chr$(DOTEL) & Chr$(ECHO)
      If TraceTelnet Then Debug.Print "SENT: DO ECHO"
      End If
    Case SGA
    If TraceTelnet Then Debug.Print "SGA"
      If Not sw_ugoahead Then
        sw_ugoahead = True
        WinsockClient.SendData Chr$(IAC) & Chr$(DOTEL) & Chr$(SGA)
      If TraceTelnet Then Debug.Print "SENT: DOTEL SGA"
      End If
    
    Case TERMSPEED
    If TraceTelnet Then Debug.Print "TERMSPEED"
    If TraceTelnet Then Debug.Print "SENT: DONT TERMSPEED"
      WinsockClient.SendData Chr$(IAC) & Chr$(DONTTEL) & Chr$(CH)
      
    Case TFLOWCNTRL
    If TraceTelnet Then Debug.Print "TFLOWCNTRL"
    If TraceTelnet Then Debug.Print "SENT: DONT FLOWCONTROL"
      WinsockClient.SendData Chr$(IAC) & Chr$(DONTTEL) & Chr$(CH)
      
    Case LINEMODE
    If TraceTelnet Then Debug.Print "LINEMODE"
    If TraceTelnet Then Debug.Print "SENT: DONT LINEMODE"
      WinsockClient.SendData Chr$(IAC) & Chr$(DONTTEL) & Chr$(CH)
      
    Case STATUS
    If TraceTelnet Then Debug.Print "STATUS"
    If TraceTelnet Then Debug.Print "SENT: DONT STATUS"
      WinsockClient.SendData Chr$(IAC) & Chr$(DONTTEL) & Chr$(CH)
      
    Case TIMING
    If TraceTelnet Then Debug.Print "TIMING"
    If TraceTelnet Then Debug.Print "SENT: DONT TIMING"
      WinsockClient.SendData Chr$(IAC) & Chr$(DONTTEL) & Chr$(CH)
      
    Case DISPLOC
    If TraceTelnet Then Debug.Print "DISPLOC"
    If TraceTelnet Then Debug.Print "SENT: WONT DISPLOC"
      WinsockClient.SendData Chr$(IAC) & Chr$(DONTTEL) & Chr$(CH)
    
    Case ENVIRON
    If TraceTelnet Then Debug.Print "ENVIRON"
    If TraceTelnet Then Debug.Print "SENT: WONT ENVIRON"
      WinsockClient.SendData Chr$(IAC) & Chr$(DONTTEL) & Chr$(CH)
    
    Case UNKNOWN39
    If TraceTelnet Then Debug.Print "UNKNOWN39"
    If TraceTelnet Then Debug.Print "SENT: WONT " & Asc(CH)
      WinsockClient.SendData Chr$(IAC) & Chr$(DONTTEL) & Chr$(CH)
    
    
    Case Else
    If TraceTelnet Then Debug.Print "UNKNOWN CMD " & Asc(CH)
    If TraceTelnet Then Debug.Print "SENT: WONT UNKNOWN CMD " & Asc(CH)
      WinsockClient.SendData Chr$(IAC) & Chr$(DONTTEL) & Chr$(CH)
  End Select

End Function

Private Function iac4(CH As Byte) As Integer

  ' WONT Processing
  
    If TraceTelnet Then Debug.Print "                                                                   RECEIVED WONT : ";

  iac4 = GO_NORM

  Select Case CH
    
    Case ECHO
    If TraceTelnet Then Debug.Print "ECHO"
      If sw_echo = True Then
        WinsockClient.SendData Chr$(IAC) & Chr$(DONTTEL) & Chr$(ECHO)
      If TraceTelnet Then Debug.Print "SENT: DONTEL ECHO"
        sw_echo = False
      End If
      
    Case SGA
    If TraceTelnet Then Debug.Print "SGA"
    If TraceTelnet Then Debug.Print "SENT: DONT SGA"
      WinsockClient.SendData Chr$(IAC) & Chr$(DONTTEL) & Chr$(SGA)
      sw_igoahead = False
    
    Case TERMSPEED
    If TraceTelnet Then Debug.Print "TERMSPEED"
    If TraceTelnet Then Debug.Print "SENT: DONT TERMSPEED"
      WinsockClient.SendData Chr$(IAC) & Chr$(DONTTEL) & Chr$(CH)
    
    Case TFLOWCNTRL
    If TraceTelnet Then Debug.Print "FLOWCONTROL"
    If TraceTelnet Then Debug.Print "SENT: DONT FLOWCONTROL"
      WinsockClient.SendData Chr$(IAC) & Chr$(DONTTEL) & Chr$(CH)
      
    Case LINEMODE
    If TraceTelnet Then Debug.Print "LINEMODE"
    If TraceTelnet Then Debug.Print "SENT: DONT LINEMODE"
      WinsockClient.SendData Chr$(IAC) & Chr$(DONTTEL) & Chr$(CH)
      
    Case STATUS
    If TraceTelnet Then Debug.Print "STATUS"
    If TraceTelnet Then Debug.Print "SENT: DONT STATUS"
      WinsockClient.SendData Chr$(IAC) & Chr$(DONTTEL) & Chr$(CH)
      
    Case TIMING
    If TraceTelnet Then Debug.Print "TIMING"
    If TraceTelnet Then Debug.Print "SENT: DONT TIMING"
      WinsockClient.SendData Chr$(IAC) & Chr$(DONTTEL) & Chr$(CH)
      
    Case DISPLOC
    If TraceTelnet Then Debug.Print "DISPLOC"
    If TraceTelnet Then Debug.Print "SENT: DONT DISPLOC"
      WinsockClient.SendData Chr$(IAC) & Chr$(DONTTEL) & Chr$(CH)
    
    Case ENVIRON
    If TraceTelnet Then Debug.Print "ENVIRON"
    If TraceTelnet Then Debug.Print "SENT: DONT ENVIRON"
      WinsockClient.SendData Chr$(IAC) & Chr$(DONTTEL) & Chr$(CH)
    
    Case UNKNOWN39
    If TraceTelnet Then Debug.Print "UNKNOWN39"
    If TraceTelnet Then Debug.Print "SENT: DONT " & Asc(CH)
      WinsockClient.SendData Chr$(IAC) & Chr$(DONTTEL) & Chr$(CH)
    
    Case Else
    If TraceTelnet Then Debug.Print "UNKNOWN CMD " & Asc(CH)
    If TraceTelnet Then Debug.Print "SENT: DONT UNKNOWN CMD " & Asc(CH)
      WinsockClient.SendData Chr$(IAC) & Chr$(WONTTEL) & Chr$(CH)
  End Select

End Function

Private Function iac5(CH As Byte) As Integer

Dim ich As Integer
  ' Collect parms after SB and until another IAC

  
    ich = CH
    If ich = IAC Then
      iac5 = GO_IAC1
      Exit Function
    End If
    
    If TraceTelnet Then Debug.Print "                                                                   RECEIVED : ";
    If TraceTelnet Then Debug.Print "SB("; ppno; ") = " & ich
    
    parsedata(ppno) = ich
    ppno = ppno + 1
    
    iac5 = GO_IAC5

End Function


Private Function iac6(CH As Byte) As Integer

  'DONT Processing

 
  iac6 = GO_NORM
        

  Select Case CH
    Case SE
      If TraceTelnet Then Debug.Print "                                                                   RECEIVED SE"
      If TraceTelnet Then Debug.Print "SENT: SE_ACK " & CH

    Case ECHO
      If TraceTelnet Then Debug.Print "                                                                   RECEIVED DONT : ";
      If TraceTelnet Then Debug.Print "ECHO"
      If Not sw_echo Then
        sw_echo = True
        WinsockClient.SendData Chr$(IAC) & Chr$(WONTTEL) & Chr$(ECHO)
        If TraceTelnet Then Debug.Print "SENT: WONT ECHO"
      End If
    Case SGA
      If TraceTelnet Then Debug.Print "                                                                   RECEIVED DONT : ";
      If TraceTelnet Then Debug.Print "SGA"
      If Not sw_ugoahead Then
        sw_ugoahead = True
        WinsockClient.SendData Chr$(IAC) & Chr$(WONTTEL) & Chr$(SGA)
        If TraceTelnet Then Debug.Print "SENT: WONT SGA"
      End If
    
    Case TERMSPEED
      If TraceTelnet Then Debug.Print "                                                                   RECEIVED DONT : ";
      If TraceTelnet Then Debug.Print "TERMSPEED"
      If TraceTelnet Then Debug.Print "SENT: WONT TERMSPEED"
      WinsockClient.SendData Chr$(IAC) & Chr$(WONTTEL) & Chr$(CH)
      
    Case TFLOWCNTRL
      If TraceTelnet Then Debug.Print "                                                                   RECEIVED DONT : ";
      If TraceTelnet Then Debug.Print "TFLOWCNTRL"
      If TraceTelnet Then Debug.Print "SENT: WONT FLOWCONTROL"
      WinsockClient.SendData Chr$(IAC) & Chr$(WONTTEL) & Chr$(CH)
      
    Case LINEMODE
      If TraceTelnet Then Debug.Print "                                                                   RECEIVED DONT : ";
      If TraceTelnet Then Debug.Print "LINEMODE"
      If TraceTelnet Then Debug.Print "SENT: WONT LINEMODE"
      WinsockClient.SendData Chr$(IAC) & Chr$(WONTTEL) & Chr$(CH)
      
    Case STATUS
      If TraceTelnet Then Debug.Print "                                                                   RECEIVED DONT : ";
      If TraceTelnet Then Debug.Print "STATUS"
      If TraceTelnet Then Debug.Print "SENT: WONT STATUS"
      WinsockClient.SendData Chr$(IAC) & Chr$(WONTTEL) & Chr$(CH)
      
    Case TIMING
      If TraceTelnet Then Debug.Print "                                                                   RECEIVED DONT : ";
      If TraceTelnet Then Debug.Print "TIMING"
      If TraceTelnet Then Debug.Print "SENT: WONT TIMING"
      WinsockClient.SendData Chr$(IAC) & Chr$(WONTTEL) & Chr$(CH)
      
    Case DISPLOC
      If TraceTelnet Then Debug.Print "                                                                   RECEIVED DONT : ";
      If TraceTelnet Then Debug.Print "DISPLOC"
      If TraceTelnet Then Debug.Print "SENT: WONT DISPLOC"
      WinsockClient.SendData Chr$(IAC) & Chr$(WONTTEL) & Chr$(CH)
    
    Case ENVIRON
      If TraceTelnet Then Debug.Print "                                                                   RECEIVED DONT : ";
      If TraceTelnet Then Debug.Print "ENVIRON"
      If TraceTelnet Then Debug.Print "SENT: WONT ENVIRON"
      WinsockClient.SendData Chr$(IAC) & Chr$(WONTTEL) & Chr$(CH)
    
    Case UNKNOWN39
      If TraceTelnet Then Debug.Print "                                                                   RECEIVED DONT : ";
      If TraceTelnet Then Debug.Print "UNKNOWN39"
      If TraceTelnet Then Debug.Print "SENT: WONT " & Asc(CH)
      WinsockClient.SendData Chr$(IAC) & Chr$(WONTTEL) & Chr$(CH)
        
    Case Else
      If TraceTelnet Then Debug.Print "                                                                   RECEIVED DONT : ";
      If TraceTelnet Then Debug.Print "UNKNOWN CMD " & Asc(CH)
      If TraceTelnet Then Debug.Print "SENT: WONT UNKNOWN CMD " & Asc(CH)
      WinsockClient.SendData Chr$(IAC) & Chr$(WONTTEL) & Chr$(CH)
  End Select

End Function


Private Sub WinsockClient_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
        
    frmTelnet.stbStatusBar.Panels(4).Text = Number & " - " & Description
    Operating = False
    Connected = False
End Sub
