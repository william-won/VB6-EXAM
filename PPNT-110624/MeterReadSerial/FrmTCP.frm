VERSION 5.00
Object = "{9F3B4DE1-AA29-11D1-A3D9-FDA4E35D1D25}#1.0#0"; "Io.ocx"
Begin VB.Form FrmSerialComm 
   BackColor       =   &H00200005&
   BorderStyle     =   0  '¾øÀ½
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   LinkTopic       =   "Form1"
   ScaleHeight     =   12053.62
   ScaleMode       =   0  '»ç¿ëÀÚ
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BackColor       =   &H00200005&
      Caption         =   "Conector Type"
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
      Height          =   1815
      Left            =   1680
      TabIndex        =   18
      Top             =   7080
      Width           =   9735
      Begin VB.TextBox txtModemIP 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
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
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   960
         Width           =   3855
      End
      Begin VB.OptionButton OptionTCP 
         BackColor       =   &H00200005&
         Caption         =   "TCP/IP Socket"
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
         Height          =   735
         Left            =   2160
         TabIndex        =   20
         Top             =   960
         Width           =   3375
      End
      Begin VB.OptionButton OptionSerial 
         BackColor       =   &H00200005&
         Caption         =   "Serial"
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
         Height          =   735
         Left            =   240
         TabIndex        =   19
         Top             =   960
         Value           =   -1  'True
         Width           =   2175
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
         Left            =   5640
         TabIndex        =   22
         Top             =   480
         Width           =   2895
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00200005&
      Caption         =   "Meter Type"
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
      Height          =   1335
      Left            =   1320
      TabIndex        =   15
      Top             =   840
      Width           =   4815
      Begin VB.OptionButton OptionGType 
         BackColor       =   &H00200005&
         Caption         =   "G-Type"
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
         Height          =   735
         Left            =   360
         TabIndex        =   17
         Top             =   480
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.OptionButton OptionEType 
         BackColor       =   &H00200005&
         Caption         =   "E-Type"
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
         Height          =   735
         Left            =   2640
         TabIndex        =   16
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.TextBox txtUsage 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
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
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   6240
      Width           =   4215
   End
   Begin VB.TextBox txtDebug 
      Height          =   1455
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   12
      Text            =   "FrmTCP.frx":0000
      Top             =   9480
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox txtTime 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
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
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   5280
      Width           =   4215
   End
   Begin VB.TextBox txtDate 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
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
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   4320
      Width           =   4215
   End
   Begin VB.TextBox txtConstant 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
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
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3360
      Width           =   4215
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
      Left            =   4800
      Style           =   1  '±×·¡ÇÈ
      TabIndex        =   3
      Top             =   9480
      Width           =   3900
   End
   Begin VB.TextBox txtMeterId 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
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
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2400
      Width           =   4215
   End
   Begin IOLib.IO IO1 
      Left            =   13560
      Top             =   720
      _Version        =   65536
      _ExtentX        =   1270
      _ExtentY        =   1270
      _StockProps     =   0
   End
   Begin VB.Label Label7 
      BackColor       =   &H00200005&
      Caption         =   "Meter Current Usage"
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
      Left            =   1800
      TabIndex        =   14
      Top             =   6240
      Width           =   5055
   End
   Begin VB.Label Label6 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
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
      Left            =   11880
      TabIndex        =   8
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackColor       =   &H00200005&
      Caption         =   "Meter Current Date"
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
      Left            =   1800
      TabIndex        =   7
      Top             =   5280
      Width           =   5055
   End
   Begin VB.Label Label4 
      BackColor       =   &H00200005&
      Caption         =   "Meter Current Time"
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
      Left            =   1800
      TabIndex        =   6
      Top             =   4440
      Width           =   5175
   End
   Begin VB.Label Label2 
      BackColor       =   &H00200005&
      Caption         =   "Meter ID"
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
      Left            =   1800
      TabIndex        =   5
      Top             =   2400
      Width           =   4335
   End
   Begin VB.Label Label1 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackColor       =   &H00200005&
      Caption         =   "Meter Test"
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
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   15135
   End
   Begin VB.Image ImgExit 
      Height          =   1980
      Left            =   13440
      Picture         =   "FrmTCP.frx":0006
      Stretch         =   -1  'True
      Top             =   9600
      Width           =   1920
   End
   Begin VB.Label Label3 
      BackColor       =   &H00200005&
      Caption         =   "Meter Constant"
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
      Left            =   1800
      TabIndex        =   2
      Top             =   3360
      Width           =   4455
   End
   Begin VB.Label txtResult 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      AutoSize        =   -1  'True
      BackColor       =   &H00200005&
      Caption         =   "O"
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
      Left            =   12090
      TabIndex        =   0
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   6120
      X2              =   0
      Y1              =   627.793
      Y2              =   627.793
   End
End
Attribute VB_Name = "FrmSerialComm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)

Dim recvlen As Integer
Dim recvbuf(200) As Byte

Dim Send_Seq_Byte As Byte

Const SNRM_Send_Len_GType = 10
Const AARQ_Send_Len_GType = 46
Const MMID_Send_Len_GType = 28
Const CMID_Send_Len_GType = 28
Const CONST_Send_Len_GType = 28
Const TIME_Send_Len_GType = 28
Const USAGE_Send_Len_GType = 28
Const DISC_Send_Len_GType = 10

Const SNRM_Send_Len_EType = 33
Const AARQ_Send_Len_EType = 75
Const MMID_Send_Len_EType = 28
Const CMID_Send_Len_EType = 28
Const CONST_Send_Len_EType = 28
Const TIME_Send_Len_EType = 28
Const USAGE_Send_Len_EType = 28
Const DISC_Send_Len_EType = 10

Public Function CalcFCS(FCS As Double, data_p() As Byte, start As Integer, length As Integer) As Long
Dim fcsTab(256) As Double
Dim i As Integer

    '0x0000, 0x1189, 0x2312, 0x329b, 0x4624, 0x57ad, 0x6536, 0x74bf,
    '0x8c48, 0x9dc1, 0xaf5a, 0xbed3, 0xca6c, 0xdbe5, 0xe97e, 0xf8f7,
    '0x1081, 0x0108, 0x3393, 0x221a, 0x56a5, 0x472c, 0x75b7, 0x643e,
    '0x9cc9, 0x8d40, 0xbfdb, 0xae52, 0xdaed, 0xcb64, 0xf9ff, 0xe876,
    '0x2102, 0x308b, 0x0210, 0x1399, 0x6726, 0x76af, 0x4434, 0x55bd,
    '0xad4a, 0xbcc3, 0x8e58, 0x9fd1, 0xeb6e, 0xfae7, 0xc87c, 0xd9f5,
    '0x3183, 0x200a, 0x1291, 0x0318, 0x77a7, 0x662e, 0x54b5, 0x453c,
    '0xbdcb, 0xac42, 0x9ed9, 0x8f50, 0xfbef, 0xea66, 0xd8fd, 0xc974,
    '0x4204, 0x538d, 0x6116, 0x709f, 0x0420, 0x15a9, 0x2732, 0x36bb,
    '0xce4c, 0xdfc5, 0xed5e, 0xfcd7, 0x8868, 0x99e1, 0xab7a, 0xbaf3,
    '0x5285, 0x430c, 0x7197, 0x601e, 0x14a1, 0x0528, 0x37b3, 0x263a,
    '0xdecd, 0xcf44, 0xfddf, 0xec56, 0x98e9, 0x8960, 0xbbfb, 0xaa72,
    '0x6306, 0x728f, 0x4014, 0x519d, 0x2522, 0x34ab, 0x0630, 0x17b9,
    '0xef4e, 0xfec7, 0xcc5c, 0xddd5, 0xa96a, 0xb8e3, 0x8a78, 0x9bf1,
    '0x7387, 0x620e, 0x5095, 0x411c, 0x35a3, 0x242a, 0x16b1, 0x0738,
    '0xffcf, 0xee46, 0xdcdd, 0xcd54, 0xb9eb, 0xa862, 0x9af9, 0x8b70,
    '0x8408, 0x9581, 0xa71a, 0xb693, 0xc22c, 0xd3a5, 0xe13e, 0xf0b7,
    '0x0840, 0x19c9, 0x2b52, 0x3adb, 0x4e64, 0x5fed, 0x6d76, 0x7cff,
    '0x9489, 0x8500, 0xb79b, 0xa612, 0xd2ad, 0xc324, 0xf1bf, 0xe036,
    '0x18c1, 0x0948, 0x3bd3, 0x2a5a, 0x5ee5, 0x4f6c, 0x7df7, 0x6c7e,
    '0xa50a, 0xb483, 0x8618, 0x9791, 0xe32e, 0xf2a7, 0xc03c, 0xd1b5,
    '0x2942, 0x38cb, 0x0a50, 0x1bd9, 0x6f66, 0x7eef, 0x4c74, 0x5dfd,
    '0xb58b, 0xa402, 0x9699, 0x8710, 0xf3af, 0xe226, 0xd0bd, 0xc134,
    '0x39c3, 0x284a, 0x1ad1, 0x0b58, 0x7fe7, 0x6e6e, 0x5cf5, 0x4d7c,
    '0xc60c, 0xd785, 0xe51e, 0xf497, 0x8028, 0x91a1, 0xa33a, 0xb2b3,
    '0x4a44, 0x5bcd, 0x6956, 0x78df, 0x0c60, 0x1de9, 0x2f72, 0x3efb,
    '0xd68d, 0xc704, 0xf59f, 0xe416, 0x90a9, 0x8120, 0xb3bb, 0xa232,
    '0x5ac5, 0x4b4c, 0x79d7, 0x685e, 0x1ce1, 0x0d68, 0x3ff3, 0x2e7a,
    '0xe70e, 0xf687, 0xc41c, 0xd595, 0xa12a, 0xb0a3, 0x8238, 0x93b1,
    '0x6b46, 0x7acf, 0x4854, 0x59dd, 0x2d62, 0x3ceb, 0x0e70, 0x1ff9,
    '0xf78f, 0xe606, 0xd49d, 0xc514, 0xb1ab, 0xa022, 0x92b9, 0x8330,
    '0x7bc7, 0x6a4e, 0x58d5, 0x495c, 0x3de3, 0x2c6a, 0x1ef1, 0x0f78

    fcsTab(0) = 0:  fcsTab(1) = 4489:   fcsTab(2) = 8978:   fcsTab(3) = 12955:  fcsTab(4) = 17956:  fcsTab(5) = 22445:  fcsTab(6) = 25910: fcsTab(7) = 29887:
    fcsTab(8) = 35912:  fcsTab(9) = 40385:  fcsTab(10) = 44890: fcsTab(11) = 48851: fcsTab(12) = 51820: fcsTab(13) = 56293: fcsTab(14) = 59774: fcsTab(15) = 63735:
    fcsTab(16) = 4225:  fcsTab(17) = 264:   fcsTab(18) = 13203: fcsTab(19) = 8730:  fcsTab(20) = 22181: fcsTab(21) = 18220: fcsTab(22) = 30135: fcsTab(23) = 25662:
    fcsTab(24) = 40137: fcsTab(25) = 36160: fcsTab(26) = 49115: fcsTab(27) = 44626: fcsTab(28) = 56045: fcsTab(29) = 52068: fcsTab(30) = 63999: fcsTab(31) = 59510:
    fcsTab(32) = 8450:  fcsTab(33) = 12427: fcsTab(34) = 528:    fcsTab(35) = 5017:  fcsTab(36) = 26406: fcsTab(37) = 30383: fcsTab(38) = 17460: fcsTab(39) = 21949:
    fcsTab(40) = 44362: fcsTab(41) = 48323: fcsTab(42) = 36440: fcsTab(43) = 40913: fcsTab(44) = 60270: fcsTab(45) = 64231: fcsTab(46) = 51324: fcsTab(47) = 55797:
    fcsTab(48) = 12675: fcsTab(49) = 8202:  fcsTab(50) = 4753:  fcsTab(51) = 792:    fcsTab(52) = 30631: fcsTab(53) = 26158: fcsTab(54) = 21685: fcsTab(55) = 17724:
    fcsTab(56) = 48587: fcsTab(57) = 44098: fcsTab(58) = 40665: fcsTab(59) = 36688: fcsTab(60) = 64495: fcsTab(61) = 60006: fcsTab(62) = 55549: fcsTab(63) = 51572:
    fcsTab(64) = 16900: fcsTab(65) = 21389: fcsTab(66) = 24854: fcsTab(67) = 28831: fcsTab(68) = 1056:   fcsTab(69) = 5545:  fcsTab(70) = 10034: fcsTab(71) = 14011:
    fcsTab(72) = 52812: fcsTab(73) = 57285: fcsTab(74) = 60766: fcsTab(75) = 64727: fcsTab(76) = 34920: fcsTab(77) = 39393: fcsTab(78) = 43898: fcsTab(79) = 47859:
    fcsTab(80) = 21125: fcsTab(81) = 17164: fcsTab(82) = 29079: fcsTab(83) = 24606: fcsTab(84) = 5281:  fcsTab(85) = 1320:   fcsTab(86) = 14259: fcsTab(87) = 9786:
    fcsTab(88) = 57037: fcsTab(89) = 53060: fcsTab(90) = 64991: fcsTab(91) = 60502: fcsTab(92) = 39145: fcsTab(93) = 35168: fcsTab(94) = 48123: fcsTab(95) = 43634
    fcsTab(96) = 25350: fcsTab(97) = 29327: fcsTab(98) = 16404: fcsTab(99) = 20893: fcsTab(100) = 9506: fcsTab(101) = 13483: fcsTab(102) = 1584: fcsTab(103) = 6073:
    fcsTab(104) = 61262: fcsTab(105) = 65223: fcsTab(106) = 52316: fcsTab(107) = 56789: fcsTab(108) = 43370: fcsTab(109) = 47331: fcsTab(110) = 35448: fcsTab(111) = 39921:
    fcsTab(112) = 29575: fcsTab(113) = 25102: fcsTab(114) = 20629: fcsTab(115) = 16668: fcsTab(116) = 13731: fcsTab(117) = 9258: fcsTab(118) = 5809: fcsTab(119) = 1848:
    fcsTab(120) = 65487: fcsTab(121) = 60998: fcsTab(122) = 56541: fcsTab(123) = 52564: fcsTab(124) = 47595: fcsTab(125) = 43106: fcsTab(126) = 39673: fcsTab(127) = 35696:
    fcsTab(128) = 33800: fcsTab(129) = 38273: fcsTab(130) = 42778: fcsTab(131) = 46739: fcsTab(132) = 49708: fcsTab(133) = 54181: fcsTab(134) = 57662: fcsTab(135) = 61623:
    fcsTab(136) = 2112: fcsTab(137) = 6601: fcsTab(138) = 11090: fcsTab(139) = 15067: fcsTab(140) = 20068: fcsTab(141) = 24557: fcsTab(142) = 28022: fcsTab(143) = 31999:
    fcsTab(144) = 38025: fcsTab(145) = 34048: fcsTab(146) = 47003: fcsTab(147) = 42514: fcsTab(148) = 53933: fcsTab(149) = 49956: fcsTab(150) = 61887: fcsTab(151) = 57398:
    fcsTab(152) = 6337: fcsTab(153) = 2376: fcsTab(154) = 15315: fcsTab(155) = 10842: fcsTab(156) = 24293: fcsTab(157) = 20332: fcsTab(158) = 32247: fcsTab(159) = 27774:
    fcsTab(160) = 42250: fcsTab(161) = 46211: fcsTab(162) = 34328: fcsTab(163) = 38801: fcsTab(164) = 58158: fcsTab(165) = 62119: fcsTab(166) = 49212: fcsTab(167) = 53685:
    fcsTab(168) = 10562: fcsTab(169) = 14539: fcsTab(170) = 2640: fcsTab(171) = 7129: fcsTab(172) = 28518: fcsTab(173) = 32495: fcsTab(174) = 19572: fcsTab(175) = 24061:
    fcsTab(176) = 46475: fcsTab(177) = 41986: fcsTab(178) = 38553: fcsTab(179) = 34576: fcsTab(180) = 62383: fcsTab(181) = 57894: fcsTab(182) = 53437: fcsTab(183) = 49460:
    fcsTab(184) = 14787: fcsTab(185) = 10314: fcsTab(186) = 6865: fcsTab(187) = 2904: fcsTab(188) = 32743: fcsTab(189) = 28270: fcsTab(190) = 23797: fcsTab(191) = 19836:
    fcsTab(192) = 50700: fcsTab(193) = 55173: fcsTab(194) = 58654: fcsTab(195) = 62615: fcsTab(196) = 32808: fcsTab(197) = 37281: fcsTab(198) = 41786: fcsTab(199) = 45747:
    fcsTab(200) = 19012: fcsTab(201) = 23501: fcsTab(202) = 26966: fcsTab(203) = 30943: fcsTab(204) = 3168: fcsTab(205) = 7657: fcsTab(206) = 12146: fcsTab(207) = 16123:
    fcsTab(208) = 54925: fcsTab(209) = 50948: fcsTab(210) = 62879: fcsTab(211) = 58390: fcsTab(212) = 37033: fcsTab(213) = 33056: fcsTab(214) = 46011: fcsTab(215) = 41522:
    fcsTab(216) = 23237: fcsTab(217) = 19276: fcsTab(218) = 31191: fcsTab(219) = 26718: fcsTab(220) = 7393: fcsTab(221) = 3432: fcsTab(222) = 16371: fcsTab(223) = 11898:
    fcsTab(224) = 59150: fcsTab(225) = 63111: fcsTab(226) = 50204: fcsTab(227) = 54677: fcsTab(228) = 41258: fcsTab(229) = 45219: fcsTab(230) = 33336: fcsTab(231) = 37809:
    fcsTab(232) = 27462: fcsTab(233) = 31439: fcsTab(234) = 18516: fcsTab(235) = 23005: fcsTab(236) = 11618: fcsTab(237) = 15595: fcsTab(238) = 3696: fcsTab(239) = 8185:
    fcsTab(240) = 63375: fcsTab(241) = 58886: fcsTab(242) = 54429: fcsTab(243) = 50452: fcsTab(244) = 45483: fcsTab(245) = 40994: fcsTab(246) = 37561: fcsTab(247) = 33584:
    fcsTab(248) = 31687: fcsTab(249) = 27214: fcsTab(250) = 22741: fcsTab(251) = 18780: fcsTab(252) = 15843: fcsTab(253) = 11370: fcsTab(254) = 7921: fcsTab(255) = 3960:
    FCS = 65535
    For i = start To (start + length - 1)
        FCS = ((FCS \ 256) And 255) Xor fcsTab((FCS Xor data_p(i)) And 255)
    Next

    FCS = FCS Xor 65535
    CalcFCS = FCS
    data_p(i) = (FCS And 255)
    data_p(i + 1) = ((FCS \ 256) And 255)

End Function
       
'=======================================================
' type conversion
'=======================================================
Public Function IntegerToUnsigned(value As Integer) As Long
    
    If value < 0 Then
        IntegerToUnsigned = value + 65536
    Else
        IntegerToUnsigned = value
    End If
    
End Function

Public Function UnsignedToInteger(value As Long) As Integer
    
    If value <= 32767 Then
        UnsignedToInteger = value
    Else
        UnsignedToInteger = value - 65536
    End If
    
End Function


' ¹ÙÀÌÆ® µ¥ÀÌÅÍ¸¦ Çí»ç string À¸·Î º¯È¯ Æã¼Ç
Public Function ByteToHexStr(bydata As Byte) As String

Dim strHex As String

strHex = Hex(bydata)

If Len(strHex) < 2 Then
strHex = "0" + strHex
End If
ByteToHexStr = strHex

End Function

Public Function ByteToChr(bydata As Byte) As String

Dim strHex As String

strHex = Chr(bydata)

ByteToChr = strHex

End Function

Private Function cmdGTypeMeterParser(Result As Byte) As Byte
Dim i As Integer
Dim ReadResult As Long
Dim startData As Byte

Dim RxString As String

Const SNRM_Recv_Len_GType = 35
Const AARE_Recv_Len_GType = 58
Const MMID_Recv_Len_GType = 28
Const CMID_Recv_Len_GType = 28
Const CONST_Recv_Len_GType = 25
Const TIME_Recv_Len_GType = 33
Const USAGE_Recv_Len_GType = 24
Const DISC_Recv_Len_GType = 10

    For i = 0 To 199: recvbuf(i) = 0: Next
    
    If IO1.NumCharsInQue = 0 Then IO1.Sleep (100)
    If IO1.NumCharsInQue = 0 Then IO1.Sleep (100)

    Select Case (Send_Seq_Byte)
    Case &H93: 'SNRM_UA
    If Result = SNRM_Send_Len_GType Then
        Result = IO1.ReadBytes(recvbuf, SNRM_Recv_Len_GType)
        If Result = 0 Or ((IO1.NumBytesRead <> SNRM_Recv_Len_GType) And (IO1.NumBytesRead <> (SNRM_Recv_Len_GType - 2))) Then
            Result = 0
        End If
    End If
    Case &H10: 'AARE
    If Result = AARQ_Send_Len_GType Then
        Result = IO1.ReadBytes(recvbuf, AARE_Recv_Len_GType)
        If Result = 0 Or (IO1.NumBytesRead <> AARE_Recv_Len_GType) Then
            Result = 0
        End If
    End If
    Case &H32 'Manufacture_Id
    If Result = MMID_Send_Len_GType Then
        For i = 0 To (MMID_Recv_Len_GType - 1)
            'IO1.Sleep (100)
            ReadResult = IO1.ReadByte()
            recvbuf(i) = ReadResult And 255
        Next
        If Result = 0 Or ((recvbuf(17) <> &H7) And (recvbuf(18) <> &H7)) Then
            Result = 0
        Else
            RxString = ""
            If (recvbuf(17) = &H7) Then
                startData = 18
            Else
                startData = 19
            End If
            
            For i = startData To (startData + 6)
                RxString = RxString + ByteToChr(recvbuf(i))
            Next i
            txtMeterId.Text = Mid(RxString, 4, 4)
        End If
    End If
    Case &H54 'Meter_Id
    If Result = CMID_Send_Len_GType Then
        For i = 0 To (CMID_Recv_Len_GType - 1)
            'IO1.Sleep (100)
            ReadResult = IO1.ReadByte()
            recvbuf(i) = ReadResult And 255
        Next
        If Result = 0 Or ((recvbuf(17) <> &H7) And (recvbuf(18) <> &H7)) Then
            Result = 0
        Else
            RxString = ""
            If (recvbuf(17) = &H7) Then
                startData = 18
            Else
                startData = 19
            End If
            
            For i = startData To (startData + 6)
                RxString = RxString + ByteToChr(recvbuf(i))
            Next i
            txtMeterId.Text = txtMeterId.Text + Mid(RxString, 1, 7) '4+7
        End If
    End If
    Case &H76 'Meter_constant
    If Result = CONST_Send_Len_GType Then
        For i = 0 To (CONST_Recv_Len_GType - 1) 'include Null data
            ReadResult = IO1.ReadByte()
            recvbuf(i) = ReadResult And 255
        Next
        If Result = 0 Or ((recvbuf(17) <> &H4) And (recvbuf(18) <> &H4)) Then
            Result = 0
        Else
            Dim f_Constant As Single
            Dim b_Constant(3) As Byte
            
            txtConstant.Text = "2000" 'float 44FA0000 => 2000
            If (recvbuf(17) = &H4) Then
                startData = 18
            Else
                startData = 19
            End If
            
            For i = startData To startData + 3
                RxString = RxString + ByteToHexStr(recvbuf(i))
                b_Constant(startData + 3 - i) = recvbuf(i)
            Next i

            Call CopyMemory(f_Constant, b_Constant(0), 4&)
            txtConstant.Text = Val(f_Constant)
            
            If txtConstant.Text = "0" Then
                Result = 0
            End If
            
        End If
    End If
    Case &H98 'Meter_Time
    If Result = TIME_Send_Len_GType Then
        'IO1.Sleep (100)
        For i = 0 To TIME_Recv_Len_GType 'include Null data
            ReadResult = IO1.ReadByte()
            recvbuf(i) = ReadResult And 255
        Next
        If Result = 0 Or ((recvbuf(17) <> &HC) And (recvbuf(18) <> &HC)) Then
            Result = 0
        Else
            RxString = ""
            If (recvbuf(17) = &HC) Then
                startData = 18
            Else
                startData = 19
            End If
            For i = startData To startData + 1
                RxString = RxString + ByteToHexStr(recvbuf(i))
            Next i

            txtDate.Text = Trim(Str("&H" + RxString))
            If Len(Trim(Str(recvbuf(startData + 2)))) = 1 Then
                txtDate.Text = txtDate.Text + ".0" + Trim(Str(recvbuf(startData + 2)))
            Else
                txtDate.Text = txtDate.Text + "." + Trim(Str(recvbuf(startData + 2)))
            End If
            If Len(Trim(Str(recvbuf(startData + 3)))) = 1 Then
                txtDate.Text = txtDate.Text + ".0" + Trim(Str(recvbuf(startData + 3)))
            Else
                txtDate.Text = txtDate.Text + "." + Trim(Str(recvbuf(startData + 3)))
            End If

            If Len(Trim(Str(recvbuf(startData + 5)))) = 1 Then
                txtTime.Text = "0" + Trim(Str(recvbuf(startData + 5)))
            Else
                txtTime.Text = Trim(Str(recvbuf(startData + 5)))
            End If
            If Len(Trim(Str(recvbuf(startData + 6)))) = 1 Then
                txtTime.Text = txtTime.Text + ":0" + Trim(Str(recvbuf(startData + 6)))
            Else
                txtTime.Text = txtTime.Text + ":" + Trim(Str(recvbuf(startData + 6)))
            End If
            If Len(Trim(Str(recvbuf(startData + 7)))) = 1 Then
                txtTime.Text = txtTime.Text + ":0" + Trim(Str(recvbuf(startData + 7)))
            Else
                txtTime.Text = txtTime.Text + ":" + Trim(Str(recvbuf(startData + 7)))
            End If
        End If
    End If
    Case &HBA 'Usage
    If Result = USAGE_Send_Len_GType Then
        For i = 0 To (USAGE_Recv_Len_GType - 1) 'include Null data
            ReadResult = IO1.ReadByte()
            recvbuf(i) = ReadResult And 255
        Next
    
        If Result = 0 Or ((recvbuf(16) <> &H6) And (recvbuf(17) <> &H6)) Then
            Result = 0
        Else
            Dim s_Usage As Single
            RxString = ""
            If (recvbuf(16) = &H6) Then
                startData = 17
            Else
                startData = 18
            End If
            For i = startData To startData + 3
                RxString = RxString + ByteToHexStr(recvbuf(i))
            Next i

            s_Usage = IntegerToUnsigned(Val("&H" + RxString)) / Val(txtConstant.Text)
            txtUsage.Text = Format(s_Usage, "##0.0") + " kWh" 'Usage/constant
        End If
    End If
    
    Case &H53 'Disc_UA
    If Result = DISC_Send_Len_GType Then
        Result = IO1.ReadBytes(recvbuf, DISC_Recv_Len_GType)
        If Result = 0 Or (IO1.NumBytesRead <> DISC_Recv_Len_GType) Then
            Result = 0
        End If
    End If
    End Select
    
    cmdGTypeMeterParser = Result
    
End Function

Private Function cmdGTypeMeterRead()
'GType Meter
Dim Result As Byte

Dim SNRM_GType(SNRM_Send_Len_GType - 1) As Byte
Dim AARQ_GType(AARQ_Send_Len_GType - 1) As Byte

Dim MANUFACTURE_ID_GType(MMID_Send_Len_GType - 1) As Byte
Dim CUSTOMER_ID_GType(CMID_Send_Len_GType - 1) As Byte
Dim CONSTANT_GType(CONST_Send_Len_GType - 1) As Byte
Dim READTIME_GType(TIME_Send_Len_GType - 1) As Byte
Dim READUSAGE_GType(USAGE_Send_Len_GType - 1) As Byte
Dim DISC_GType(DISC_Send_Len_GType - 1) As Byte

Dim HCS As Double
Dim FCS As Double

    SNRM_GType(0) = &H7E: SNRM_GType(1) = &HA0: SNRM_GType(2) = &H8: SNRM_GType(3) = &H2: SNRM_GType(4) = &H3
    SNRM_GType(5) = &H21: SNRM_GType(6) = &H93: SNRM_GType(7) = &H86: SNRM_GType(8) = &H67: SNRM_GType(9) = &H7E
    
    AARQ_GType(0) = &H7E: AARQ_GType(1) = &HA0: AARQ_GType(2) = &H2C: AARQ_GType(3) = &H2: AARQ_GType(4) = &H3
    AARQ_GType(5) = &H21: AARQ_GType(6) = &H10: AARQ_GType(7) = &H94: AARQ_GType(8) = &H9C: AARQ_GType(9) = &HE6
    AARQ_GType(10) = &HE6: AARQ_GType(11) = &H0: AARQ_GType(12) = &H60: AARQ_GType(13) = &H1D: AARQ_GType(14) = &HA1
    AARQ_GType(15) = &H9: AARQ_GType(16) = &H6: AARQ_GType(17) = &H7: AARQ_GType(18) = &H60: AARQ_GType(19) = &H85
    AARQ_GType(20) = &H74: AARQ_GType(21) = &H5: AARQ_GType(22) = &H8: AARQ_GType(23) = &H1: AARQ_GType(24) = &H1
    AARQ_GType(25) = &HBE: AARQ_GType(26) = &H10: AARQ_GType(27) = &H4: AARQ_GType(28) = &HE: AARQ_GType(29) = &H1
    AARQ_GType(30) = &H0: AARQ_GType(31) = &H0: AARQ_GType(32) = &H0: AARQ_GType(33) = &H6: AARQ_GType(34) = &H5F
    AARQ_GType(35) = &H1F: AARQ_GType(36) = &H4: AARQ_GType(37) = &H0: AARQ_GType(38) = &H0: AARQ_GType(39) = &H18
    AARQ_GType(40) = &H19: AARQ_GType(41) = &HFF: AARQ_GType(42) = &HFF: AARQ_GType(43) = &H3E: AARQ_GType(44) = &HCC
    AARQ_GType(45) = &H7E

    MANUFACTURE_ID_GType(0) = &H7E: MANUFACTURE_ID_GType(1) = &HA0: MANUFACTURE_ID_GType(2) = &H1A: MANUFACTURE_ID_GType(3) = &H2: MANUFACTURE_ID_GType(4) = &H3
    MANUFACTURE_ID_GType(5) = &H23: MANUFACTURE_ID_GType(6) = &H32: MANUFACTURE_ID_GType(7) = &H7D: MANUFACTURE_ID_GType(8) = &H42: MANUFACTURE_ID_GType(9) = &HE6
    MANUFACTURE_ID_GType(10) = &HE6: MANUFACTURE_ID_GType(11) = &H0: MANUFACTURE_ID_GType(12) = &HC0: MANUFACTURE_ID_GType(13) = &H1: MANUFACTURE_ID_GType(14) = &H81
    MANUFACTURE_ID_GType(15) = &H0: MANUFACTURE_ID_GType(16) = &H1: MANUFACTURE_ID_GType(17) = &H1: MANUFACTURE_ID_GType(18) = &H0: MANUFACTURE_ID_GType(19) = &H0
    MANUFACTURE_ID_GType(20) = &H0: MANUFACTURE_ID_GType(21) = &H1: MANUFACTURE_ID_GType(22) = &HFF: MANUFACTURE_ID_GType(23) = &H2: MANUFACTURE_ID_GType(24) = &H0
    MANUFACTURE_ID_GType(25) = &H7D: MANUFACTURE_ID_GType(26) = &H7C: MANUFACTURE_ID_GType(27) = &H7E:

    CUSTOMER_ID_GType(0) = &H7E: CUSTOMER_ID_GType(1) = &HA0: CUSTOMER_ID_GType(2) = &H1A: CUSTOMER_ID_GType(3) = &H2: CUSTOMER_ID_GType(4) = &H3
    CUSTOMER_ID_GType(5) = &H23: CUSTOMER_ID_GType(6) = &H54: CUSTOMER_ID_GType(7) = &H4D: CUSTOMER_ID_GType(8) = &H44: CUSTOMER_ID_GType(9) = &HE6
    CUSTOMER_ID_GType(10) = &HE6: CUSTOMER_ID_GType(11) = &H0: CUSTOMER_ID_GType(12) = &HC0: CUSTOMER_ID_GType(13) = &H1: CUSTOMER_ID_GType(14) = &H81
    CUSTOMER_ID_GType(15) = &H0: CUSTOMER_ID_GType(16) = &H1: CUSTOMER_ID_GType(17) = &H1: CUSTOMER_ID_GType(18) = &H0: CUSTOMER_ID_GType(19) = &H0
    CUSTOMER_ID_GType(20) = &H0: CUSTOMER_ID_GType(21) = &H0: CUSTOMER_ID_GType(22) = &HFF: CUSTOMER_ID_GType(23) = &H2: CUSTOMER_ID_GType(24) = &H0
    CUSTOMER_ID_GType(25) = &HC6: CUSTOMER_ID_GType(26) = &H60: CUSTOMER_ID_GType(27) = &H7E

    CONSTANT_GType(0) = &H7E:   CONSTANT_GType(1) = &HA0:    CONSTANT_GType(2) = &H1A:    CONSTANT_GType(3) = &H2:     CONSTANT_GType(4) = &H3:
    CONSTANT_GType(5) = &H23:    CONSTANT_GType(6) = &H76:    CONSTANT_GType(7) = &HCA:    CONSTANT_GType(8) = &H6F:    CONSTANT_GType(9) = &HE6
    CONSTANT_GType(10) = &HE6:    CONSTANT_GType(11) = &H0:     CONSTANT_GType(12) = &HC0:    CONSTANT_GType(13) = &H1:     CONSTANT_GType(14) = &H81
    CONSTANT_GType(15) = &H0:     CONSTANT_GType(16) = &H3:     CONSTANT_GType(17) = &H1:     CONSTANT_GType(18) = &H1:     CONSTANT_GType(19) = &H0
    CONSTANT_GType(20) = &H3:     CONSTANT_GType(21) = &H0:     CONSTANT_GType(22) = &HFF:    CONSTANT_GType(23) = &H2:     CONSTANT_GType(24) = &H0
    CONSTANT_GType(25) = &H25:    CONSTANT_GType(26) = &H79:    CONSTANT_GType(27) = &H7E

    READTIME_GType(0) = &H7E:   READTIME_GType(1) = &HA0:    READTIME_GType(2) = &H1A:    READTIME_GType(3) = &H2:     READTIME_GType(4) = &H3
    READTIME_GType(5) = &H23:    READTIME_GType(6) = &H98:    READTIME_GType(7) = &HBA:    READTIME_GType(8) = &H61:    READTIME_GType(9) = &HE6
    READTIME_GType(10) = &HE6:    READTIME_GType(11) = &H0:     READTIME_GType(12) = &HC0:    READTIME_GType(13) = &H1:     READTIME_GType(14) = &H81
    READTIME_GType(15) = &H0:     READTIME_GType(16) = &H8:     READTIME_GType(17) = &H0:     READTIME_GType(18) = &H0:     READTIME_GType(19) = &H1
    READTIME_GType(20) = &H0:     READTIME_GType(21) = &H0:     READTIME_GType(22) = &HFF:    READTIME_GType(23) = &H2:     READTIME_GType(24) = &H0
    READTIME_GType(25) = &H65:    READTIME_GType(26) = &HD7:    READTIME_GType(27) = &H7E

    READUSAGE_GType(0) = &H7E:   READUSAGE_GType(1) = &HA0:    READUSAGE_GType(2) = &H1A:    READUSAGE_GType(3) = &H2:     READUSAGE_GType(4) = &H3
    READUSAGE_GType(5) = &H23:    READUSAGE_GType(6) = &HBA:    READUSAGE_GType(7) = &HAA:    READUSAGE_GType(8) = &H63:    READUSAGE_GType(9) = &HE6
    READUSAGE_GType(10) = &HE6:    READUSAGE_GType(11) = &H0:     READUSAGE_GType(12) = &HC0:    READUSAGE_GType(13) = &H1:     READUSAGE_GType(14) = &H81
    READUSAGE_GType(15) = &H0:     READUSAGE_GType(16) = &H3:     READUSAGE_GType(17) = &H1:     READUSAGE_GType(18) = &H1:     READUSAGE_GType(19) = &H1
    READUSAGE_GType(20) = &H8:     READUSAGE_GType(21) = &H0:     READUSAGE_GType(22) = &HFF:    READUSAGE_GType(23) = &H2:     READUSAGE_GType(24) = &H0
    READUSAGE_GType(25) = &HE2:    READUSAGE_GType(26) = &H3A:    READUSAGE_GType(27) = &H7E

    DISC_GType(0) = &H7E:   DISC_GType(1) = &HA0:    DISC_GType(2) = &H8:     DISC_GType(3) = &H2:     DISC_GType(4) = &H3
    DISC_GType(5) = &H21:    DISC_GType(6) = &H53:    DISC_GType(7) = &H8A:    DISC_GType(8) = &HA1:    DISC_GType(9) = &H7E
    
txtResult.Caption = ""
InitView

Result = True

'IO1.Sleep (100)
If Result > 0 Then
    If (OptionSerial.value) Then
        Result = IO1.Open("COM1:", "baud=9600 parity=N data=8 stop=1")
    Else
        Result = IO1.Open(txtModemIP + ":40044", "client") 'IO1.Open("192.168.1.5:9092", "client") Open a TCP Port 9092.
    End If
    IO1.SetTimeOut (100)
    IO1.Mode = 0 '2: Async Mode
End If

If Result > 0 Then
    If IO1.NumCharsInQue <> 0 Then Result = IO1.ReadBytes(recvbuf, IO1.NumCharsInQue) 'discard data
    HCS = CalcFCS(HCS, SNRM_GType, 1, 6) 'HCS
    Send_Seq_Byte = &H93
    SNRM_GType(6) = Send_Seq_Byte
    Result = IO1.WriteBytes(SNRM_GType)
    While IO1.NumCharsOutQue <> 0
    Wend
    Result = cmdGTypeMeterParser(Result)
End If

'IO1.Sleep (100)
If Result > 0 Then
    If IO1.NumCharsInQue <> 0 Then Result = IO1.ReadBytes(recvbuf, IO1.NumCharsInQue)
    HCS = CalcFCS(HCS, AARQ_GType, 1, 6) 'HCS
    FCS = CalcFCS(HCS, AARQ_GType, 1, (AARQ_Send_Len_GType - 4)) 'fcs len-4
    Send_Seq_Byte = &H10
    AARQ_GType(6) = Send_Seq_Byte
    Result = IO1.WriteBytes(AARQ_GType)
    While IO1.NumCharsOutQue <> 0
    Wend
    Result = cmdGTypeMeterParser(Result)
End If

'IO1.Sleep (100)
If Result > 0 Then
    If IO1.NumCharsInQue <> 0 Then Result = IO1.ReadBytes(recvbuf, IO1.NumCharsInQue)
    HCS = CalcFCS(HCS, MANUFACTURE_ID_GType, 1, 6) 'HCS
    FCS = CalcFCS(HCS, MANUFACTURE_ID_GType, 1, (MMID_Send_Len_GType - 4)) 'fcs len-4
    Send_Seq_Byte = (Send_Seq_Byte + &H22) And 255
    MANUFACTURE_ID_GType(6) = Send_Seq_Byte
    Result = IO1.WriteBytes(MANUFACTURE_ID_GType)
    While IO1.NumCharsOutQue <> 0
    Wend
    Result = cmdGTypeMeterParser(Result)
End If

'IO1.Sleep (100)
If Result > 0 Then
    If IO1.NumCharsInQue <> 0 Then Result = IO1.ReadBytes(recvbuf, IO1.NumCharsInQue)
    HCS = CalcFCS(HCS, CUSTOMER_ID_GType, 1, 6) 'HCS
    FCS = CalcFCS(HCS, CUSTOMER_ID_GType, 1, (CMID_Send_Len_GType - 4)) 'fcs len-4
    Send_Seq_Byte = (Send_Seq_Byte + &H22) And 255
    CUSTOMER_ID_GType(6) = Send_Seq_Byte
    Result = IO1.WriteBytes(CUSTOMER_ID_GType)
    While IO1.NumCharsOutQue <> 0
    Wend
    Result = cmdGTypeMeterParser(Result)
End If

'IO1.Sleep (100)
If Result > 0 Then
    If IO1.NumCharsInQue <> 0 Then Result = IO1.ReadBytes(recvbuf, IO1.NumCharsInQue)
    HCS = CalcFCS(HCS, CONSTANT_GType, 1, 6) 'HCS
    FCS = CalcFCS(HCS, CONSTANT_GType, 1, (CONST_Send_Len_GType - 4)) 'fcs len-4
    Send_Seq_Byte = (Send_Seq_Byte + &H22) And 255
    CONSTANT_GType(6) = Send_Seq_Byte
    Result = IO1.WriteBytes(CONSTANT_GType)
    While IO1.NumCharsOutQue <> 0
    Wend
    Result = cmdGTypeMeterParser(Result)
End If

'IO1.Sleep (100)
If Result > 0 Then
    If IO1.NumCharsInQue <> 0 Then Result = IO1.ReadBytes(recvbuf, IO1.NumCharsInQue)
    HCS = CalcFCS(HCS, READTIME_GType, 1, 6) 'HCS
    FCS = CalcFCS(HCS, READTIME_GType, 1, (TIME_Send_Len_GType - 4)) 'fcs len-4
    Send_Seq_Byte = (Send_Seq_Byte + &H22) And 255
    READTIME_GType(6) = Send_Seq_Byte
    Result = IO1.WriteBytes(READTIME_GType) 'yy yy mm dd 0xff hh mm ss
    While IO1.NumCharsOutQue <> 0
    Wend
    Result = cmdGTypeMeterParser(Result)
End If

'IO1.Sleep (100)
If Result > 0 Then
    If IO1.NumCharsInQue <> 0 Then Result = IO1.ReadBytes(recvbuf, IO1.NumCharsInQue)
    HCS = CalcFCS(HCS, READUSAGE_GType, 1, 6) 'HCS
    FCS = CalcFCS(HCS, READUSAGE_GType, 1, (USAGE_Send_Len_GType - 4)) 'fcs len-4
    Send_Seq_Byte = (Send_Seq_Byte + &H22) And 255
    READUSAGE_GType(6) = Send_Seq_Byte
    Result = IO1.WriteBytes(READUSAGE_GType)
    While IO1.NumCharsOutQue <> 0
    Wend
    Result = cmdGTypeMeterParser(Result)
End If

'IO1.Sleep (100)
If Result > 0 Then
    HCS = CalcFCS(HCS, DISC_GType, 1, 6) 'HCS
    Send_Seq_Byte = &H53
    DISC_GType(6) = Send_Seq_Byte
    Result = IO1.WriteBytes(DISC_GType)
    While IO1.NumCharsOutQue <> 0
    Wend
    Result = cmdGTypeMeterParser(Result)
Else 'Always Send Disc
    HCS = CalcFCS(HCS, DISC_GType, 1, 6) 'HCS
    Send_Seq_Byte = &H53
    DISC_GType(6) = Send_Seq_Byte
    Result = IO1.WriteBytes(DISC_GType)
    While IO1.NumCharsOutQue <> 0
    Wend
    Result = cmdGTypeMeterParser(Result)
    Result = 0
End If

IO1.Close

If Result > 0 Then
    txtResult.ForeColor = &HFF00& '(&H0000FF00& = Light Green)
    txtResult.Caption = "O"
Else
    txtResult.ForeColor = &HFF& '(&H000000FF& = Red)
    txtResult.Caption = "X"
End If

End Function

Private Function cmdETypeMeterParser(Result As Byte) As Byte
Dim i As Integer
Dim ReadResult As Long
Dim startData As Byte

Dim RxString As String

Const SNRM_Recv_Len_EType = 35
Const AARE_Recv_Len_EType = 58
Const MMID_Recv_Len_EType = 28
Const CMID_Recv_Len_EType = 28
Const CONST_Recv_Len_EType = 25
Const TIME_Recv_Len_EType = 33
Const USAGE_Recv_Len_EType = 24
Const DISC_Recv_Len_EType = 10

    For i = 0 To 199: recvbuf(i) = 0: Next
    
    If IO1.NumCharsInQue = 0 Then IO1.Sleep (100)
    If IO1.NumCharsInQue = 0 Then IO1.Sleep (100)

    Select Case (Send_Seq_Byte)
    Case &H93: 'SNRM_UA
    If Result = SNRM_Send_Len_EType Then
        Result = IO1.ReadBytes(recvbuf, SNRM_Recv_Len_EType)
        If Result = 0 Or ((IO1.NumBytesRead <> SNRM_Recv_Len_EType) And (IO1.NumBytesRead <> (SNRM_Recv_Len_EType - 2))) Then
            Result = 0
        End If
    End If
    Case &H10: 'AARE
    If Result = AARQ_Send_Len_EType Then
        Result = IO1.ReadBytes(recvbuf, AARE_Recv_Len_EType)
        If Result = 0 Or (IO1.NumBytesRead <> AARE_Recv_Len_EType) Then
            Result = 0
        End If
    End If
    Case &H32 'Manufacture_Id
    If Result = MMID_Send_Len_EType Then
        For i = 0 To (MMID_Recv_Len_EType - 1)
            'IO1.Sleep (100)
            ReadResult = IO1.ReadByte()
            recvbuf(i) = ReadResult And 255
        Next
        If Result = 0 Or ((recvbuf(17) <> &H7) And (recvbuf(18) <> &H7)) Then
            Result = 0
        Else
            RxString = ""
            If (recvbuf(17) = &H7) Then
                startData = 18
            Else
                startData = 19
            End If
            
            For i = startData To (startData + 6)
                RxString = RxString + ByteToChr(recvbuf(i))
            Next i
            txtMeterId.Text = Mid(RxString, 4, 4)
        End If
    End If
    Case &H54 'Meter_Id
    If Result = CMID_Send_Len_EType Then
        For i = 0 To (CMID_Recv_Len_EType - 1)
            'IO1.Sleep (100)
            ReadResult = IO1.ReadByte()
            recvbuf(i) = ReadResult And 255
        Next
        If Result = 0 Or ((recvbuf(17) <> &H7) And (recvbuf(18) <> &H7)) Then
            Result = 0
        Else
            RxString = ""
            If (recvbuf(17) = &H7) Then
                startData = 18
            Else
                startData = 19
            End If
            
            For i = startData To (startData + 6)
                RxString = RxString + ByteToChr(recvbuf(i))
            Next i
            txtMeterId.Text = txtMeterId.Text + Mid(RxString, 1, 7) '4+7
        End If
    End If
    Case &H76 'Meter_constant
    If Result = CONST_Send_Len_EType Then
        For i = 0 To (CONST_Recv_Len_EType - 1) 'include Null data
            ReadResult = IO1.ReadByte()
            recvbuf(i) = ReadResult And 255
        Next
        If Result = 0 Or ((recvbuf(14) <> &H81) And (recvbuf(15) <> &H81)) Then
            Result = 0
        Else
            Dim f_Constant As Single
            Dim b_Constant(3) As Byte
            txtConstant.Text = "2000" 'float 44FA0000 => 2000
            If (recvbuf(14) = &H81) Then
                startData = 17
            Else
                startData = 18
            End If
            
            For i = startData To startData + 3
                RxString = RxString + ByteToHexStr(recvbuf(i))
                b_Constant(startData + 3 - i) = recvbuf(i)
            Next i

            Call CopyMemory(f_Constant, b_Constant(0), 4&)
            txtConstant.Text = Val(f_Constant)
            
            If txtConstant.Text = "0" Then
                Result = 0
            End If
        End If
    End If
    Case &H98 'Meter_Time
    If Result = TIME_Send_Len_EType Then
        'IO1.Sleep (100)
        For i = 0 To TIME_Recv_Len_EType 'include Null data
            ReadResult = IO1.ReadByte()
            recvbuf(i) = ReadResult And 255
        Next
        If Result = 0 Or ((recvbuf(17) <> &HC) And (recvbuf(18) <> &HC)) Then
            Result = 0
        Else
            RxString = ""
            If (recvbuf(17) = &HC) Then
                startData = 18
            Else
                startData = 19
            End If
            For i = startData To startData + 1
                RxString = RxString + ByteToHexStr(recvbuf(i))
            Next i

            txtDate.Text = Trim(Str("&H" + RxString))
            If Len(Trim(Str(recvbuf(startData + 2)))) = 1 Then
                txtDate.Text = txtDate.Text + ".0" + Trim(Str(recvbuf(startData + 2)))
            Else
                txtDate.Text = txtDate.Text + "." + Trim(Str(recvbuf(startData + 2)))
            End If
            If Len(Trim(Str(recvbuf(startData + 3)))) = 1 Then
                txtDate.Text = txtDate.Text + ".0" + Trim(Str(recvbuf(startData + 3)))
            Else
                txtDate.Text = txtDate.Text + "." + Trim(Str(recvbuf(startData + 3)))
            End If

            If Len(Trim(Str(recvbuf(startData + 5)))) = 1 Then
                txtTime.Text = "0" + Trim(Str(recvbuf(startData + 5)))
            Else
                txtTime.Text = Trim(Str(recvbuf(startData + 5)))
            End If
            If Len(Trim(Str(recvbuf(startData + 6)))) = 1 Then
                txtTime.Text = txtTime.Text + ":0" + Trim(Str(recvbuf(startData + 6)))
            Else
                txtTime.Text = txtTime.Text + ":" + Trim(Str(recvbuf(startData + 6)))
            End If
            If Len(Trim(Str(recvbuf(startData + 7)))) = 1 Then
                txtTime.Text = txtTime.Text + ":0" + Trim(Str(recvbuf(startData + 7)))
            Else
                txtTime.Text = txtTime.Text + ":" + Trim(Str(recvbuf(startData + 7)))
            End If
        End If
    End If
    Case &HBA 'Usage
    If Result = USAGE_Send_Len_EType Then
        For i = 0 To (USAGE_Recv_Len_EType - 1) 'include Null data
            ReadResult = IO1.ReadByte()
            recvbuf(i) = ReadResult And 255
        Next
    
        If Result = 0 Or ((recvbuf(16) <> &H6) And (recvbuf(17) <> &H6)) Then
            Result = 0
        Else
            Dim s_Usage As Single
            RxString = ""
            If (recvbuf(16) = &H6) Then
                startData = 17
            Else
                startData = 18
            End If
            For i = startData To startData + 3
                RxString = RxString + ByteToHexStr(recvbuf(i))
            Next i

            s_Usage = IntegerToUnsigned(Val("&H" + RxString)) / Val(txtConstant.Text)
            txtUsage.Text = Format(s_Usage, "##0.0") + " kWh" 'Usage/constant
        End If
    End If
    
    Case &H53 'Disc_UA
    If Result = DISC_Send_Len_EType Then
        Result = IO1.ReadBytes(recvbuf, DISC_Recv_Len_EType)
        If Result = 0 Or (IO1.NumBytesRead <> DISC_Recv_Len_EType) Then
            Result = 0
        End If
    End If
    End Select
    
    cmdETypeMeterParser = Result
    
End Function

Private Function cmdEtypeMeterRead()
'Etype Meter
Dim Result As Byte
Dim i As Byte

Dim SNRM_EType(SNRM_Send_Len_EType - 1) As Byte
Dim AARQ_EType(AARQ_Send_Len_EType - 1) As Byte

Dim MANUFACTURE_ID_EType(MMID_Send_Len_EType - 1) As Byte
Dim CUSTOMER_ID_EType(CMID_Send_Len_EType - 1) As Byte
Dim CONSTANT_EType(CONST_Send_Len_EType - 1) As Byte
Dim READTIME_EType(TIME_Send_Len_EType - 1) As Byte
Dim READUSAGE_EType(USAGE_Send_Len_EType - 1) As Byte
Dim DISC_EType(DISC_Send_Len_EType - 1) As Byte

Dim RxString As String
Dim ReadResult As Long

Dim HCS As Double
Dim FCS As Double

    SNRM_EType(0) = &H7E: SNRM_EType(1) = &HA0: SNRM_EType(2) = &H1F: SNRM_EType(3) = &H2:  SNRM_EType(4) = &HFF: SNRM_EType(5) = &H23: SNRM_EType(6) = &H93: SNRM_EType(7) = &H3D
    SNRM_EType(8) = &HF9: SNRM_EType(9) = &H81: SNRM_EType(10) = &H80: SNRM_EType(11) = &H12: SNRM_EType(12) = &H5: SNRM_EType(13) = &H1: SNRM_EType(14) = &H82: SNRM_EType(15) = &H6
    SNRM_EType(16) = &H1: SNRM_EType(17) = &H82: SNRM_EType(18) = &H7: SNRM_EType(19) = &H4: SNRM_EType(20) = &H0: SNRM_EType(21) = &H0: SNRM_EType(22) = &H0: SNRM_EType(23) = &H2
    SNRM_EType(24) = &H8: SNRM_EType(25) = &H4: SNRM_EType(26) = &H0: SNRM_EType(27) = &H0: SNRM_EType(28) = &H0: SNRM_EType(29) = &H2: SNRM_EType(30) = &HCD: SNRM_EType(31) = &HBE
    SNRM_EType(32) = &H7E
    
    AARQ_EType(0) = &H7E:  AARQ_EType(1) = &HA0:  AARQ_EType(2) = &H49:  AARQ_EType(3) = &H2:   AARQ_EType(4) = &HFF
    AARQ_EType(5) = &H23:  AARQ_EType(6) = &H10:  AARQ_EType(7) = &H54:  AARQ_EType(8) = &H1:   AARQ_EType(9) = &HE6
    AARQ_EType(10) = &HE6:  AARQ_EType(11) = &H0:   AARQ_EType(12) = &H60:  AARQ_EType(13) = &H3A:  AARQ_EType(14) = &H80
    AARQ_EType(15) = &H2:   AARQ_EType(16) = &H2:   AARQ_EType(17) = &H84:  AARQ_EType(18) = &HA1:  AARQ_EType(19) = &H9
    AARQ_EType(20) = &H6:   AARQ_EType(21) = &H7:   AARQ_EType(22) = &H60:  AARQ_EType(23) = &H85:  AARQ_EType(24) = &H74
    AARQ_EType(25) = &H5:   AARQ_EType(26) = &H8:   AARQ_EType(27) = &H1:   AARQ_EType(28) = &H1:   AARQ_EType(29) = &H8A
    AARQ_EType(30) = &H2:   AARQ_EType(31) = &H7:   AARQ_EType(32) = &H80:  AARQ_EType(33) = &H8B:  AARQ_EType(34) = &H7
    AARQ_EType(35) = &H60:  AARQ_EType(36) = &H85:  AARQ_EType(37) = &H74:  AARQ_EType(38) = &H5:   AARQ_EType(39) = &H8
    AARQ_EType(40) = &H2:   AARQ_EType(41) = &H1:   AARQ_EType(42) = &HAC:  AARQ_EType(43) = &HA:   AARQ_EType(44) = &H80
    AARQ_EType(45) = &H8:   AARQ_EType(46) = &H31:  AARQ_EType(47) = &H41:  AARQ_EType(48) = &H32:  AARQ_EType(49) = &H42
    AARQ_EType(50) = &H33:  AARQ_EType(51) = &H43:  AARQ_EType(52) = &H34:  AARQ_EType(53) = &H44:  AARQ_EType(54) = &HBE
    AARQ_EType(55) = &H10:  AARQ_EType(56) = &H4:   AARQ_EType(57) = &HE:   AARQ_EType(58) = &H1:   AARQ_EType(59) = &H0
    AARQ_EType(60) = &H0:   AARQ_EType(61) = &H0:   AARQ_EType(62) = &H6:   AARQ_EType(63) = &H5F:  AARQ_EType(64) = &H1F
    AARQ_EType(65) = &H4:   AARQ_EType(66) = &H0:   AARQ_EType(67) = &H0:   AARQ_EType(68) = &H18:  AARQ_EType(69) = &H1D
    AARQ_EType(70) = &H4:   AARQ_EType(71) = &H0:   AARQ_EType(72) = &H91:  AARQ_EType(73) = &H84:  AARQ_EType(74) = &H7E

    MANUFACTURE_ID_EType(0) = &H7E:    MANUFACTURE_ID_EType(1) = &HA0:    MANUFACTURE_ID_EType(2) = &H1A:    MANUFACTURE_ID_EType(3) = &H2
    MANUFACTURE_ID_EType(4) = &HFF:    MANUFACTURE_ID_EType(5) = &H23:    MANUFACTURE_ID_EType(6) = &H32:    MANUFACTURE_ID_EType(7) = &H54
    MANUFACTURE_ID_EType(8) = &H1:     MANUFACTURE_ID_EType(9) = &HE6:    MANUFACTURE_ID_EType(10) = &HE6:    MANUFACTURE_ID_EType(11) = &H0
    MANUFACTURE_ID_EType(12) = &HC0:    MANUFACTURE_ID_EType(13) = &H1:     MANUFACTURE_ID_EType(14) = &H81:    MANUFACTURE_ID_EType(15) = &H0
    MANUFACTURE_ID_EType(16) = &H1:     MANUFACTURE_ID_EType(17) = &H1:     MANUFACTURE_ID_EType(18) = &H0:     MANUFACTURE_ID_EType(19) = &H0
    MANUFACTURE_ID_EType(20) = &H0:     MANUFACTURE_ID_EType(21) = &H1:     MANUFACTURE_ID_EType(22) = &HFF:    MANUFACTURE_ID_EType(23) = &H2
    MANUFACTURE_ID_EType(24) = &H0:     MANUFACTURE_ID_EType(25) = &H7D:    MANUFACTURE_ID_EType(26) = &H7C:    MANUFACTURE_ID_EType(27) = &H7E
    
    CUSTOMER_ID_EType(0) = &H7E:   CUSTOMER_ID_EType(1) = &HA0:    CUSTOMER_ID_EType(2) = &H1A:    CUSTOMER_ID_EType(3) = &H2:     CUSTOMER_ID_EType(4) = &HFF
    CUSTOMER_ID_EType(5) = &H23:    CUSTOMER_ID_EType(6) = &H54:    CUSTOMER_ID_EType(7) = &HDA:    CUSTOMER_ID_EType(8) = &H6D:    CUSTOMER_ID_EType(9) = &HE6
    CUSTOMER_ID_EType(10) = &HE6:    CUSTOMER_ID_EType(11) = &H0:     CUSTOMER_ID_EType(12) = &HC0:    CUSTOMER_ID_EType(13) = &H1:    CUSTOMER_ID_EType(14) = &H81
    CUSTOMER_ID_EType(15) = &H0:     CUSTOMER_ID_EType(16) = &H1:     CUSTOMER_ID_EType(17) = &H1:    CUSTOMER_ID_EType(18) = &H0:     CUSTOMER_ID_EType(19) = &H0
    CUSTOMER_ID_EType(20) = &H0:     CUSTOMER_ID_EType(21) = &H0:    CUSTOMER_ID_EType(22) = &HFF:    CUSTOMER_ID_EType(23) = &H2:     CUSTOMER_ID_EType(24) = &H0
    CUSTOMER_ID_EType(25) = &HC6:    CUSTOMER_ID_EType(26) = &H60:    CUSTOMER_ID_EType(27) = &H7E

    CONSTANT_EType(0) = &H7E:   CONSTANT_EType(1) = &HA0:    CONSTANT_EType(2) = &H1A:    CONSTANT_EType(3) = &H2:     CONSTANT_EType(4) = &HFF:
    CONSTANT_EType(5) = &H23:    CONSTANT_EType(6) = &H76:    CONSTANT_EType(7) = &HCA:    CONSTANT_EType(8) = &H6F:    CONSTANT_EType(9) = &HE6
    CONSTANT_EType(10) = &HE6:    CONSTANT_EType(11) = &H0:     CONSTANT_EType(12) = &HC0:    CONSTANT_EType(13) = &H1:     CONSTANT_EType(14) = &H81
    CONSTANT_EType(15) = &H0:     CONSTANT_EType(16) = &H3:     CONSTANT_EType(17) = &H1:     CONSTANT_EType(18) = &H1:     CONSTANT_EType(19) = &H0
    CONSTANT_EType(20) = &H3:     CONSTANT_EType(21) = &H0:     CONSTANT_EType(22) = &HFF:    CONSTANT_EType(23) = &H2:     CONSTANT_EType(24) = &H0
    CONSTANT_EType(25) = &H25:    CONSTANT_EType(26) = &H79:    CONSTANT_EType(27) = &H7E
    
    READTIME_EType(0) = &H7E:   READTIME_EType(1) = &HA0:    READTIME_EType(2) = &H1A:    READTIME_EType(3) = &H2:     READTIME_EType(4) = &HFF
    READTIME_EType(5) = &H23:    READTIME_EType(6) = &H98:    READTIME_EType(7) = &HBA:    READTIME_EType(8) = &H61:    READTIME_EType(9) = &HE6
    READTIME_EType(10) = &HE6:    READTIME_EType(11) = &H0:     READTIME_EType(12) = &HC0:    READTIME_EType(13) = &H1:     READTIME_EType(14) = &H81
    READTIME_EType(15) = &H0:     READTIME_EType(16) = &H8:     READTIME_EType(17) = &H0:     READTIME_EType(18) = &H0:     READTIME_EType(19) = &H1
    READTIME_EType(20) = &H0:     READTIME_EType(21) = &H0:     READTIME_EType(22) = &HFF:    READTIME_EType(23) = &H2:     READTIME_EType(24) = &H0
    READTIME_EType(25) = &H65:    READTIME_EType(26) = &HD7:    READTIME_EType(27) = &H7E
    
    READUSAGE_EType(0) = &H7E:   READUSAGE_EType(1) = &HA0:    READUSAGE_EType(2) = &H1A:    READUSAGE_EType(3) = &H2:     READUSAGE_EType(4) = &HFF
    READUSAGE_EType(5) = &H23:    READUSAGE_EType(6) = &HBA:    READUSAGE_EType(7) = &HAA:    READUSAGE_EType(8) = &H63:    READUSAGE_EType(9) = &HE6
    READUSAGE_EType(10) = &HE6:    READUSAGE_EType(11) = &H0:     READUSAGE_EType(12) = &HC0:    READUSAGE_EType(13) = &H1:     READUSAGE_EType(14) = &H81
    READUSAGE_EType(15) = &H0:     READUSAGE_EType(16) = &H3:     READUSAGE_EType(17) = &H1:     READUSAGE_EType(18) = &H1:     READUSAGE_EType(19) = &H1
    READUSAGE_EType(20) = &H8:     READUSAGE_EType(21) = &H0:     READUSAGE_EType(22) = &HFF:    READUSAGE_EType(23) = &H2:     READUSAGE_EType(24) = &H0
    READUSAGE_EType(25) = &HE2:    READUSAGE_EType(26) = &H3A:    READUSAGE_EType(27) = &H7E

    DISC_EType(0) = &H7E:   DISC_EType(1) = &HA0:    DISC_EType(2) = &H8:     DISC_EType(3) = &H2:     DISC_EType(4) = &HFF
    DISC_EType(5) = &H23:    DISC_EType(6) = &H53:    DISC_EType(7) = &HAD:    DISC_EType(8) = &HBB:    DISC_EType(9) = &H7E

txtResult.Caption = ""
InitView

Result = True

'IO1.Sleep (100)
If Result > 0 Then
    If (OptionSerial.value) Then
        Result = IO1.Open("COM1:", "baud=9600 parity=N data=8 stop=1")
    Else
        Result = IO1.Open(txtModemIP + ":40044", "client") 'IO1.Open("192.168.1.5:9092", "client") Open a TCP Port 9092.
    End If
    IO1.SetTimeOut (100)
    IO1.Mode = 0 '2: Async Mode
End If

If Result > 0 Then
    If IO1.NumCharsInQue <> 0 Then Result = IO1.ReadBytes(recvbuf, IO1.NumCharsInQue) 'discard data
    HCS = CalcFCS(HCS, SNRM_EType, 1, 6) 'HCS
    Send_Seq_Byte = &H93
    SNRM_EType(6) = Send_Seq_Byte
    Result = IO1.WriteBytes(SNRM_EType)
    While IO1.NumCharsOutQue <> 0
    Wend
    Result = cmdETypeMeterParser(Result)
End If

'IO1.Sleep (100)
If Result > 0 Then
    If IO1.NumCharsInQue <> 0 Then Result = IO1.ReadBytes(recvbuf, IO1.NumCharsInQue)
    HCS = CalcFCS(HCS, AARQ_EType, 1, 6) 'HCS
    FCS = CalcFCS(HCS, AARQ_EType, 1, (AARQ_Send_Len_EType - 4)) 'fcs len-4
    Send_Seq_Byte = &H10
    AARQ_EType(6) = Send_Seq_Byte
    Result = IO1.WriteBytes(AARQ_EType)
    While IO1.NumCharsOutQue <> 0
    Wend
    Result = cmdETypeMeterParser(Result)
End If

'IO1.Sleep (100)
If Result > 0 Then
    If IO1.NumCharsInQue <> 0 Then Result = IO1.ReadBytes(recvbuf, IO1.NumCharsInQue)
    HCS = CalcFCS(HCS, MANUFACTURE_ID_EType, 1, 6) 'HCS
    FCS = CalcFCS(HCS, MANUFACTURE_ID_EType, 1, (MMID_Send_Len_EType - 4)) 'fcs len-4
    Send_Seq_Byte = (Send_Seq_Byte + &H22) And 255
    MANUFACTURE_ID_EType(6) = Send_Seq_Byte
    Result = IO1.WriteBytes(MANUFACTURE_ID_EType)
    While IO1.NumCharsOutQue <> 0
    Wend
    Result = cmdETypeMeterParser(Result)
End If

'IO1.Sleep (100)
If Result > 0 Then
    If IO1.NumCharsInQue <> 0 Then Result = IO1.ReadBytes(recvbuf, IO1.NumCharsInQue)
    HCS = CalcFCS(HCS, CUSTOMER_ID_EType, 1, 6) 'HCS
    FCS = CalcFCS(HCS, CUSTOMER_ID_EType, 1, (CMID_Send_Len_EType - 4)) 'fcs len-4
    Send_Seq_Byte = (Send_Seq_Byte + &H22) And 255
    CUSTOMER_ID_EType(6) = Send_Seq_Byte
    Result = IO1.WriteBytes(CUSTOMER_ID_EType)
    While IO1.NumCharsOutQue <> 0
    Wend
    Result = cmdETypeMeterParser(Result)
End If

'IO1.Sleep (100)
If Result > 0 Then
    If IO1.NumCharsInQue <> 0 Then Result = IO1.ReadBytes(recvbuf, IO1.NumCharsInQue)
    HCS = CalcFCS(HCS, CONSTANT_EType, 1, 6) 'HCS
    FCS = CalcFCS(HCS, CONSTANT_EType, 1, (CONST_Send_Len_EType - 4)) 'fcs len-4
    Send_Seq_Byte = (Send_Seq_Byte + &H22) And 255
    CONSTANT_EType(6) = Send_Seq_Byte
    Result = IO1.WriteBytes(CONSTANT_EType)
    While IO1.NumCharsOutQue <> 0
    Wend
    Result = cmdETypeMeterParser(Result)
End If

'IO1.Sleep (100)
If Result > 0 Then
    If IO1.NumCharsInQue <> 0 Then Result = IO1.ReadBytes(recvbuf, IO1.NumCharsInQue)
    HCS = CalcFCS(HCS, READTIME_EType, 1, 6) 'HCS
    FCS = CalcFCS(HCS, READTIME_EType, 1, (TIME_Send_Len_EType - 4)) 'fcs len-4
    Send_Seq_Byte = (Send_Seq_Byte + &H22) And 255
    READTIME_EType(6) = Send_Seq_Byte
    Result = IO1.WriteBytes(READTIME_EType) 'yy yy mm dd 0xff hh mm ss
    While IO1.NumCharsOutQue <> 0
    Wend
    Result = cmdETypeMeterParser(Result)
End If

'IO1.Sleep (100)
If Result > 0 Then
    If IO1.NumCharsInQue <> 0 Then Result = IO1.ReadBytes(recvbuf, IO1.NumCharsInQue)
    HCS = CalcFCS(HCS, READUSAGE_EType, 1, 6) 'HCS
    FCS = CalcFCS(HCS, READUSAGE_EType, 1, (USAGE_Send_Len_EType - 4)) 'fcs len-4
    Send_Seq_Byte = (Send_Seq_Byte + &H22) And 255
    READUSAGE_EType(6) = Send_Seq_Byte
    Result = IO1.WriteBytes(READUSAGE_EType)
    While IO1.NumCharsOutQue <> 0
    Wend
    Result = cmdETypeMeterParser(Result)
End If

'IO1.Sleep (100)
If Result > 0 Then
    HCS = CalcFCS(HCS, DISC_EType, 1, 6) 'HCS
    Send_Seq_Byte = &H53
    DISC_EType(6) = Send_Seq_Byte
    Result = IO1.WriteBytes(DISC_EType)
    While IO1.NumCharsOutQue <> 0
    Wend
    Result = cmdETypeMeterParser(Result)
Else 'Always Send Disc
    HCS = CalcFCS(HCS, DISC_EType, 1, 6) 'HCS
    Send_Seq_Byte = &H53
    DISC_EType(6) = Send_Seq_Byte
    Result = IO1.WriteBytes(DISC_EType)
    While IO1.NumCharsOutQue <> 0
    Wend
    Result = cmdETypeMeterParser(Result)
    Result = 0
End If

IO1.Close

If Result > 0 Then
    txtResult.ForeColor = &HFF00& '(&H0000FF00& = Light Green)
    txtResult.Caption = "O"
Else
    txtResult.ForeColor = &HFF& '(&H000000FF& = Red)
    txtResult.Caption = "X"
End If

End Function

Private Sub cmdStart_Click()
    InitView
    If (OptionGType.value) Then
        cmdGTypeMeterRead
    Else
        cmdEtypeMeterRead
    End If
End Sub

Private Sub Form_Load()

Frame1.Visible = False
txtModemIP.Enabled = False
If (0) Then
    Label1.Caption = "G-TypeMeter Tester"
    OptionGType.value = True
Else
    Label1.Caption = "E-Type Meter Tester"
    OptionEType.value = True
End If
    InitView
End Sub

Public Function InitView()
    txtMeterId.Text = ""
    txtConstant.Text = ""
    txtDate.Text = ""
    txtTime.Text = ""
    txtUsage.Text = ""
    Send_Seq_Byte = 0
    txtResult.Caption = ""
End Function

Private Sub ImgExit_Click()
Unload Me
End Sub

Private Sub OptionEType_Click()
    Label1.Caption = "E-Type Meter Tester"
End Sub

Private Sub OptionGType_Click()
    Label1.Caption = "G-Type Meter Tester"
End Sub

Private Sub OptionSerial_Click()
    txtModemIP.Enabled = False
End Sub

Private Sub OptionTCP_Click()
    txtModemIP.Enabled = True
End Sub
