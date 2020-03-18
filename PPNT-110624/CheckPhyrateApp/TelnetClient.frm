VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTelnet 
   BackColor       =   &H00200005&
   BorderStyle     =   0  '없음
   Caption         =   "CheckPhyRateApp v1.0     Copyright (c) ATAW"
   ClientHeight    =   11010
   ClientLeft      =   2475
   ClientTop       =   1995
   ClientWidth     =   15240
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
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   ShowInTaskbar   =   0   'False
   Begin VB.Timer TimerState 
      Interval        =   3000
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00200005&
      Caption         =   "Modem Infomation"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   5895
      Left            =   720
      TabIndex        =   7
      Top             =   1200
      Width           =   10575
      Begin VB.CommandButton cmdModify 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Modify"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   7560
         Style           =   1  '그래픽
         TabIndex        =   22
         Top             =   4800
         Width           =   2460
      End
      Begin VB.TextBox txtMode 
         Alignment       =   2  '가운데 맞춤
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Text            =   "-"
         Top             =   5040
         Width           =   855
      End
      Begin VB.TextBox txtSubmode 
         Alignment       =   2  '가운데 맞춤
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Text            =   "-"
         Top             =   5040
         Width           =   855
      End
      Begin VB.TextBox txtVersion 
         Alignment       =   2  '가운데 맞춤
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Text            =   "-"
         Top             =   1200
         Width           =   5295
      End
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
         Left            =   3720
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   "-"
         Top             =   3000
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
         Left            =   600
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "-"
         Top             =   3000
         Width           =   1815
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
         Left            =   8040
         TabIndex        =   10
         Text            =   "5"
         Top             =   3360
         Width           =   855
      End
      Begin VB.TextBox txtMAC 
         Alignment       =   2  '가운데 맞춤
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   600
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   "-"
         Top             =   1200
         Width           =   3615
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         X1              =   0
         X2              =   10560
         Y1              =   4680
         Y2              =   4680
      End
      Begin VB.Label Label6 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00200005&
         Caption         =   "Mode"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   495
         Left            =   600
         TabIndex        =   21
         Top             =   5160
         Width           =   1215
      End
      Begin VB.Label Label9 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00200005&
         Caption         =   "SubMode"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   495
         Left            =   3840
         TabIndex        =   20
         Top             =   5160
         Width           =   1935
      End
      Begin VB.Label Label10 
         BackColor       =   &H00200005&
         Caption         =   "F/W Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   495
         Left            =   6360
         TabIndex        =   17
         Top             =   480
         Width           =   3255
      End
      Begin VB.Label Label2 
         BackColor       =   &H00200005&
         Caption         =   "RX PhyRate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   615
         Left            =   3360
         TabIndex        =   15
         Top             =   2040
         Width           =   2895
      End
      Begin VB.Label Label3 
         BackColor       =   &H00200005&
         Caption         =   "Tx PhyRate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   615
         Left            =   120
         TabIndex        =   14
         Top             =   2040
         Width           =   2775
      End
      Begin VB.Label Label4 
         BackColor       =   &H00200005&
         Caption         =   "Low Rate Alarm"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   735
         Left            =   6600
         TabIndex        =   13
         Top             =   2040
         Width           =   3735
      End
      Begin VB.Label Label5 
         BackColor       =   &H00200005&
         Caption         =   "MAC Address"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   495
         Left            =   1080
         TabIndex        =   9
         Top             =   600
         Width           =   3255
      End
   End
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4680
      Style           =   1  '그래픽
      TabIndex        =   3
      Top             =   9120
      Width           =   3900
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00200005&
      Caption         =   "Modem Parameter"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   1695
      Left            =   720
      TabIndex        =   0
      Top             =   7320
      Width           =   10575
      Begin VB.TextBox txtRemoteAddress 
         Alignment       =   2  '가운데 맞춤
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   5640
         MaxLength       =   15
         TabIndex        =   1
         TabStop         =   0   'False
         Text            =   "192.168.0.106"
         Top             =   720
         Width           =   3855
      End
      Begin VB.Label Label8 
         BackColor       =   &H00200005&
         Caption         =   "IP Address"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   495
         Left            =   480
         TabIndex        =   2
         Top             =   840
         Width           =   2895
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   11880
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock WinsockClient 
      Left            =   12840
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label txtResult 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      BackColor       =   &H00200005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   99.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   2295
      Left            =   12375
      TabIndex        =   6
      Top             =   3240
      Width           =   585
   End
   Begin VB.Image ImgExit 
      Height          =   1980
      Left            =   13320
      Picture         =   "TelnetClient.frx":0442
      Stretch         =   -1  'True
      Top             =   9240
      Width           =   1920
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00200005&
      Caption         =   "Marvell Modem Test"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   1
      Left            =   0
      TabIndex        =   5
      Top             =   120
      Width           =   15135
   End
   Begin VB.Label Label7 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00200005&
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   735
      Left            =   11760
      TabIndex        =   4
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   10080
      Y1              =   0
      Y2              =   0
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
Private DoDiscAfterComplete As Boolean
Private CheckModeDS2 As Boolean

Private readInfo As Boolean
Private readSubmode As Boolean
Private readVersion As Boolean
Private WaitAnswer As Boolean
Private SendNext As Integer
'Private IsMaster As Boolean
Private setMode As Boolean
Private setSubMode As Boolean
Private setFixedNM1 As Boolean
Private setFixedNM2 As Boolean

Private Function cmdinfo_Click()
  If Connected Then
    Dim CH As String
    
    If readInfo = True Then
        CH = "/i" & vbCrLf
    ElseIf readSubmode = True Then
        CH = "/tm gs" & vbCrLf
    ElseIf readVersion = True Then
        CH = "/ver" & vbCrLf
    Else
        CH = "/i" & vbCrLf
    End If
    Debug.Print "Tx:" & CH
    WinsockClient.SendData CH
  Else
    'NotConnectedAlarm
    txtResult.Caption = "X"
    txtResult.ForeColor = &HFF&
  End If
End Function

Private Function cmdSet()

  If Connected Then
    Dim CH As String
    Dim data As Long
    Dim txtSend As String

    If setSubMode = True Then
        CH = "/fs write nvram 0x2c10 0x" & txtSubmode.Text & vbCrLf
        Debug.Print "Tx:" & CH
        WinsockClient.SendData CH
        
        '다른 값을 입력하더라도 nvram과 동일한 값 입력 후 submode 변경시
        '192.168.0.106/255.255.0.0/192.168.0.1/10/4로 초기화됨
        setFixedNM1 = True
        setFixedNM2 = True
    ElseIf setFixedNM1 = True Then
        CH = "/n ip nm 255 0 0 0" & vbCrLf
        Debug.Print "Tx:" & CH
        WinsockClient.SendData CH
    ElseIf setFixedNM2 = True Then
        CH = "/n ip nm 255 255 0 0" & vbCrLf
        Debug.Print "Tx:" & CH
        WinsockClient.SendData CH
    '모드 변경 시 직접 접속이 아니라 PLC 연결 통해서 모뎀에 연결 된 경우 연결이 끊기므로 마지막에 변경한다.
    ElseIf setMode = True Then
        CH = "/s m w " & txtMode.Text & vbCrLf
        Debug.Print "Tx:" & CH
        WinsockClient.SendData CH
    Else '명령어 완료
        CH = "/hw rst" & vbCrLf
        Debug.Print "Tx:" & CH
        WinsockClient.SendData CH
        WaitAnswer = False

        Shell ("arp -d")
        Shell ("arp -d")
        txtResult.Caption = "O"
        txtResult.ForeColor = &HFF00&
        DoDiscAfterComplete = True

    End If
  Else
    'NotConnectedAlarm
    txtResult.Caption = "X"
    txtResult.ForeColor = &HFF&
  End If
End Function


Private Sub Connect_Click()
   On Error Resume Next                                   ' Handle errors...
'------------------------------------------------------------
    If Not Operating Then
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
            frmTelnet.RemotePort = "40000"
            frmTelnet.WinsockClient.RemotePort = "40000"
            frmTelnet.WinsockClient.RemoteHost = txtRemoteAddress
            .RemoteHost = RemoteIPAd 'txtRemoteAddress
            .RemotePort = RemotePort 'txtPort
            .Connect ' Attempt new connection
            'term_init
        End With
    End If

End Sub

Private Sub Disconnect_Click()
    WinsockClient_Close
    txtMAC.Text = "-"
    txtMode.Text = "-"
    txtTxPhyRate.Text = "-"
    txtRxPhyRate.Text = "-"
    txtResult.Caption = ""
    txtResult.ForeColor = &H80000005
End Sub

Private Function NotConnectedAlarm()
  MsgBox "연결되어있지않음", , "알림"
End Function

Private Sub cmdModify_Click()

Dim checkerr As Boolean
Dim txtlen As Integer
Dim txtcheck As String
checkerr = True


If (checkerr = True) And ((IsNumeric(txtMode) = False) Or (IsNumeric(txtSubmode) = False)) Then
    checkerr = False
End If

  If checkerr = False Then
    MsgBox "입력값 오류", , "알림"
    txtResult.Caption = "X"
    txtResult.ForeColor = &HFF&
  ElseIf Connected Then
    WaitAnswer = True
              
    setMode = True
    setSubMode = True
    txtResult.Caption = "-"
    txtResult.ForeColor = &HFFFF&
    Call cmdSet
  Else
    'NotConnectedAlarm
    txtResult.Caption = "X"
    txtResult.ForeColor = &HFF&
  End If
End Sub

Private Sub cmdStart_Click()
    Connect_Click
End Sub


Private Sub Form_Load()
    Dim i As Integer
    
    Me.Top = 0
    Me.Left = 0

    TimerState.Enabled = True
    
    RemoteIPAd = "192.168.0.106"

    WaitAnswer = False
    CheckModeDS2 = True
    readSubmode = False

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

Private Sub ImgExit_Click()
Unload Me
End Sub

Private Sub TimerState_Timer()
    If (WaitAnswer = False) Then
        If Connected Then
            Call cmdinfo_Click
        End If
    End If
End Sub

Private Sub WinsockClient_sendcomplete()
    Sendcomplete = True

    If DoDiscAfterComplete Then
      DoDiscAfterComplete = False
      Call WinsockClient_Close
    End If
End Sub


Private Function initView()
        txtMAC.Text = "-"
        txtMode.Text = "-"
        txtSubmode.Text = "-"
        txtVersion.Text = "-"
        txtTxPhyRate.Text = "-"
        txtRxPhyRate.Text = "-"
        txtResult.Caption = ""
        txtResult.ForeColor = &H80000005
        setMode = False
        setSubMode = False
        setFixedNM1 = False
        setFixedNM2 = False

        readSubmode = False
        readVersion = False
        'IsMaster = False
End Function

Private Sub WinsockClient_Close()
        
      If TraceTelnet Then Debug.Print Int(Timer) & " - [Closed  ] : Connection Reset By Peer "
        With WinsockClient
            .Close                                     ' Clear any errors...
            .RemotePort = 0
            .LocalPort = 0
        End With
        Operating = False
        Connected = False
        TimerState.Interval = 150 'read Version, MAC, Mode, SubMode, att timeout
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
        initView
        Connected = True
        readInfo = True 'read once
        
        If CheckModeDS2 Then
          Dim CH As String

          CH = "mode ds2" & vbCrLf
          Debug.Print "Tx:" & CH
          WinsockClient.SendData CH
        End If
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

Private Sub WinsockClient_DataArrival(ByVal bytesTotal As Long)
  Dim rxStr As String
  Dim disStr As String
  Dim cChar As String
  Dim i As Integer
  Dim j As Integer
  Dim CH As String
  Dim MACAddress As String

  WinsockClient.GetData rxStr

  If InStr(rxStr, "#user@/>") > 0 And CheckModeDS2 Then
    CH = "/mode ds2" & vbCrLf
    Debug.Print "Tx:" & CH
    WinsockClient.SendData CH
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
          If cChar = "" Then
          '
          ElseIf (AscB(cChar) = 255) And (StrComp(Mid(rxStr, i + 1, 1), "?") = 0) Then
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

  If InStr(rxStr, "Password:") > 0 Then
    If CheckModeDS2 Then
      CH = "frigodedo" & vbCrLf
    End If
    Debug.Print "Tx:" & CH
    WinsockClient.SendData CH
  End If
  
  If (bytesTotal > 200) Then
      If InStr(rxStr, "MAC: 00:") > 0 Then
        j = InStr(rxStr, "MODE: ") + 5
        txtMode = Trim(Mid(rxStr, j, 3))
    
        j = InStr(rxStr, "MAC: 00:") + 5
        txtMAC = Trim(Mid(rxStr, j, 17))
        
        If InStr(rxStr, "Forwarding (M)") > 0 Then
            '11. 00:0B:29:00:0A:B2     36 Mbps        33 Mbps       Forwarding (M)
            ' 9. 00:07:7F:88:01:57     125 Mbps        10 Mbps       Forwarding (M)
            j = InStr(rxStr, "Forwarding (M)")
        
            txtTxPhyRate = Trim(Mid(rxStr, j - 30, 3))
            txtRxPhyRate = Trim(Mid(rxStr, j - 15, 3))
        End If
        If ((Val(txtTxPhyRate.Text) <= Val(txtRawRate.Text)) Or (Val(txtRxPhyRate.Text) <= Val(txtRawRate.Text))) Then
            txtResult = "X"
            txtResult.ForeColor = &HFF&
        Else
            txtResult = "O"
            txtResult.ForeColor = &HFF00&
        End If
        'If InStr(rxStr, "Master Access") > 0 Then
          'IsMaster = True
        'End If
        
        If readInfo = True Then
          readInfo = False
          readSubmode = True 'read Next
        End If
      End If
  End If

'Hardware MAC Address: 00:07:7F:00:00:21
'IP Address          : 192.168.1.182
'Subnet Mask         : 255.255.0.0
'Gateway Address     : 192.168.0.1
  If (readSubmode = True) And (InStr(rxStr, "Current SUB-MODE is ") > 0) Then
    txtSubmode.Text = Trim(Mid(rxStr, InStr(rxStr, "is ") + 3, 1))
    readSubmode = False
    readVersion = True 'read Next
  ElseIf (readVersion = True) And (InStr(rxStr, "FIRMWARE VERSION") > 0) Then
    txtVersion.Text = Trim(Mid(rxStr, Len("FIRMWARE VERSION") + 2, InStr(rxStr, "running") - Len("FIRMWARE VERSION") - 3))
    readVersion = False
    TimerState.Interval = 0
  ElseIf (bytesTotal < 100) And (InStr(rxStr, "OK") > 0) Then 'MAC Input OK response 67 Byte, IP OK Res 73 Byte
    If (WaitAnswer = True) Then
      Dim txtSend As String
      
      If setMode = False And setSubMode = False And setFixedNM1 = False And setFixedNM2 = False Then
        CH = "/hw rst" & vbCrLf
        Debug.Print "Tx:" & CH
        WinsockClient.SendData CH
        WaitAnswer = False

        Shell ("arp -d")
        Shell ("arp -d")
        txtResult.Caption = "O"
        txtResult.ForeColor = &HFF00&
        DoDiscAfterComplete = True
      Else
        
        If setSubMode = True Then '순서 주의!! Submode->NM->GW->Mode
            setSubMode = False
        ElseIf setFixedNM1 = True Then
            setFixedNM1 = False
        ElseIf setFixedNM2 = True Then
            setFixedNM2 = False
        ElseIf setMode = True Then '모드 변경은 맨 마지막
            setMode = False
        End If
        Call cmdSet
      End If
    Else
        Call cmdinfo_Click
    End If
  ElseIf (InStr(rxStr, "KO") > 0) Then
   If setMode = True And (InStr(rxStr, "NOT VALID") > 0) Then
    txtResult.ForeColor = "X"
    txtResult.ForeColor = &HFF&
    WaitAnswer = False
   End If
  End If
  
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
    Operating = False
    Connected = False
End Sub
