VERSION 5.00
Begin VB.Form frmSetParameters 
   BackColor       =   &H00800000&
   Caption         =   "Parameters"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   6585
   StartUpPosition =   3  'Windows ±âº»°ª
   Begin VB.CommandButton cmdSet 
      Caption         =   "Set"
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
      Left            =   960
      TabIndex        =   18
      Top             =   4200
      Width           =   2175
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
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
      Left            =   4200
      TabIndex        =   17
      Top             =   4200
      Width           =   2175
   End
   Begin VB.ComboBox ComboCur 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2520
      TabIndex        =   16
      Text            =   "Combo1"
      Top             =   2280
      Width           =   1455
   End
   Begin VB.ComboBox ComboAttn 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2520
      TabIndex        =   15
      Text            =   "Combo1"
      Top             =   3360
      Width           =   1455
   End
   Begin VB.ComboBox ComboVpp 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   480
      TabIndex        =   14
      Text            =   "Combo1"
      Top             =   3360
      Width           =   1455
   End
   Begin VB.OptionButton OptionOff 
      BackColor       =   &H00800000&
      Caption         =   "Off"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   4320
      TabIndex        =   12
      Top             =   240
      Width           =   975
   End
   Begin VB.OptionButton OptionOn 
      BackColor       =   &H00800000&
      Caption         =   "On"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   3480
      TabIndex        =   10
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Text            =   "3"
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Text            =   "1000"
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   1
      Text            =   "12"
      Top             =   1200
      Width           =   1455
   End
   Begin VB.ComboBox ComboBand 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   480
      TabIndex        =   0
      Text            =   "C Band"
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackColor       =   &H00800000&
      Caption         =   "Tx CurLimit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2160
      TabIndex        =   13
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Åõ¸í
      Caption         =   "CENELEC Protocol"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   480
      TabIndex        =   11
      Top             =   240
      Width           =   2895
   End
   Begin VB.Label Label8 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackColor       =   &H00800000&
      Caption         =   "Attn"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2640
      TabIndex        =   9
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackColor       =   &H00800000&
      BackStyle       =   0  'Åõ¸í
      Caption         =   "TxVpp"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   600
      TabIndex        =   8
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label9 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackColor       =   &H00800000&
      Caption         =   "MAX Retry"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label10 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackColor       =   &H00800000&
      Caption         =   "Send Pkts"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label13 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackColor       =   &H00800000&
      Caption         =   "BAND"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   720
      TabIndex        =   5
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label14 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackColor       =   &H00800000&
      Caption         =   "Pkt Size"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4680
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "frmSetParameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public MaxRetries As Integer


Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdSet_Click()
MaxRetries = Val(Text6.Text)
Unload Me
End Sub

Private Sub Form_Load()
ComboVpp.Clear
ComboVpp.AddItem "1.7 V"
ComboVpp.AddItem "3.5 V"
ComboVpp.AddItem "7 V"
ComboVpp.AddItem "15 V"
ComboVpp.ListIndex = 2

ComboAttn.Clear
ComboAttn.AddItem "0 dB"
ComboAttn.AddItem "6 dB"
ComboAttn.AddItem "12 dB"
ComboAttn.AddItem "18 dB"
ComboAttn.AddItem "24 dB"
ComboAttn.ListIndex = 2

ComboBand.Clear
ComboBand.AddItem "A Band"
ComboBand.AddItem "B Band"
ComboBand.AddItem "C Band"
ComboBand.ListIndex = 2

ComboCur.Clear
ComboCur.AddItem "1 A"
ComboCur.AddItem "2 A"
ComboCur.ListIndex = 0

OptionOn.Value = True

End Sub
