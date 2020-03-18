VERSION 5.00
Object = "{A8B345A0-74B5-11D3-85C2-00105AC8B715}#1.0#0"; "iProfessionalLibrary.ocx"
Begin VB.Form frmSend 
   BackColor       =   &H00200005&
   BorderStyle     =   0  '¾øÀ½
   Caption         =   "Form1"
   ClientHeight    =   4005
   ClientLeft      =   1995
   ClientTop       =   4995
   ClientWidth     =   7845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   7845
   ShowInTaskbar   =   0   'False
   Begin iProfessionalLibrary.iSwitchRocker3WayX iSwitchRocker3WayX1 
      Height          =   1215
      Left            =   6000
      TabIndex        =   2
      Top             =   1080
      Width           =   735
      ShowFocusRect   =   -1  'True
      BorderMargin    =   2
      RepeatInitialDelay=   500
      RepeatInterval  =   50
      Value           =   0
      Increment       =   1
      UseArrowKeys    =   -1  'True
      BackGroundColor =   2097157
      BorderStyle     =   0
      Enabled         =   -1  'True
      Transparent     =   0   'False
      UpdateFrameRate =   60
      OptionSaveAllProperties=   0   'False
      AutoFrameRate   =   0   'False
      Object.Visible         =   -1  'True
      Object.Width           =   49
      Object.Height          =   81
      OPCItemCount    =   0
   End
   Begin VB.TextBox txpowertxt 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      TabIndex        =   1
      Text            =   "0"
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackColor       =   &H00200005&
      Caption         =   "Down"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   6960
      TabIndex        =   4
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label18 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackColor       =   &H00200005&
      Caption         =   "Up"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   6960
      TabIndex        =   3
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00200005&
      Caption         =   "Tx Power"
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
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   1440
      Width           =   2295
   End
End
Attribute VB_Name = "frmSend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer
Private Sub iSwitchRocker3WayX1_OnValueChange()
    a = iSwitchRocker3WayX1.Value
    If a < 0 Then
    txpowertxt.Text = 0
    iSwitchRocker3WayX1.Value = 0
    MsgBox "¹üÀ§¸¦ ¹þ¾î³³´Ï´Ù.", vbExclamation, "Setting"
    Else
    txpowertxt.Text = a
    End If
End Sub
