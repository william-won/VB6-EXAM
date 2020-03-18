VERSION 5.00
Object = "{A8B345A0-74B5-11D3-85C2-00105AC8B715}#1.0#0"; "iProfessionalLibrary.ocx"
Begin VB.Form frmSetting 
   BackColor       =   &H00200005&
   BorderStyle     =   0  '없음
   Caption         =   "Form1"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSetting.frx":0000
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   Begin iProfessionalLibrary.iSwitchRocker3WayX iSwitchRocker3WayX1 
      Height          =   1335
      Left            =   9000
      TabIndex        =   11
      Top             =   2520
      Width           =   1095
      ShowFocusRect   =   -1  'True
      BorderMargin    =   2
      RepeatInitialDelay=   500
      RepeatInterval  =   50
      Value           =   0
      Increment       =   1
      UseArrowKeys    =   -1  'True
      BackGroundColor =   16761024
      BorderStyle     =   0
      Enabled         =   -1  'True
      Transparent     =   0   'False
      UpdateFrameRate =   60
      OptionSaveAllProperties=   0   'False
      AutoFrameRate   =   0   'False
      Object.Visible         =   -1  'True
      Object.Width           =   73
      Object.Height          =   89
      OPCItemCount    =   0
   End
   Begin VB.TextBox Amptxt 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00200005&
      BorderStyle     =   0  '없음
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
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "0"
      Top             =   7440
      Width           =   1095
   End
   Begin VB.TextBox StopFtxt 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00200005&
      BorderStyle     =   0  '없음
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
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "0"
      Top             =   6000
      Width           =   1095
   End
   Begin VB.TextBox StartBtxt 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00200005&
      BorderStyle     =   0  '없음
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
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "0"
      Top             =   4440
      Width           =   1095
   End
   Begin VB.TextBox SampFtxt 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00200005&
      BorderStyle     =   0  '없음
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
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "0"
      Top             =   3000
      Width           =   1095
   End
   Begin iProfessionalLibrary.iSwitchRocker3WayX iSwitchRocker3WayX2 
      Height          =   1335
      Left            =   11040
      TabIndex        =   12
      Top             =   3960
      Width           =   1095
      ShowFocusRect   =   -1  'True
      BorderMargin    =   2
      RepeatInitialDelay=   500
      RepeatInterval  =   50
      Value           =   0
      Increment       =   10
      UseArrowKeys    =   -1  'True
      BackGroundColor =   16761024
      BorderStyle     =   0
      Enabled         =   -1  'True
      Transparent     =   0   'False
      UpdateFrameRate =   60
      OptionSaveAllProperties=   0   'False
      AutoFrameRate   =   0   'False
      Object.Visible         =   -1  'True
      Object.Width           =   73
      Object.Height          =   89
      OPCItemCount    =   0
   End
   Begin iProfessionalLibrary.iSwitchRocker3WayX iSwitchRocker3WayX3 
      Height          =   1335
      Left            =   11040
      TabIndex        =   13
      Top             =   5520
      Width           =   1095
      ShowFocusRect   =   -1  'True
      BorderMargin    =   2
      RepeatInitialDelay=   500
      RepeatInterval  =   50
      Value           =   0
      Increment       =   10
      UseArrowKeys    =   -1  'True
      BackGroundColor =   16761024
      BorderStyle     =   0
      Enabled         =   -1  'True
      Transparent     =   0   'False
      UpdateFrameRate =   60
      OptionSaveAllProperties=   0   'False
      AutoFrameRate   =   0   'False
      Object.Visible         =   -1  'True
      Object.Width           =   73
      Object.Height          =   89
      OPCItemCount    =   0
   End
   Begin iProfessionalLibrary.iSwitchRocker3WayX iSwitchRocker3WayX4 
      Height          =   1335
      Left            =   11040
      TabIndex        =   14
      Top             =   2520
      Width           =   1095
      ShowFocusRect   =   -1  'True
      BorderMargin    =   2
      RepeatInitialDelay=   500
      RepeatInterval  =   50
      Value           =   0
      Increment       =   10
      UseArrowKeys    =   -1  'True
      BackGroundColor =   16761024
      BorderStyle     =   0
      Enabled         =   -1  'True
      Transparent     =   0   'False
      UpdateFrameRate =   60
      OptionSaveAllProperties=   0   'False
      AutoFrameRate   =   0   'False
      Object.Visible         =   -1  'True
      Object.Width           =   73
      Object.Height          =   89
      OPCItemCount    =   0
   End
   Begin iProfessionalLibrary.iSwitchRocker3WayX iSwitchRocker3WayX5 
      Height          =   1335
      Left            =   9000
      TabIndex        =   15
      Top             =   3960
      Width           =   1095
      ShowFocusRect   =   -1  'True
      BorderMargin    =   2
      RepeatInitialDelay=   500
      RepeatInterval  =   50
      Value           =   0
      Increment       =   1
      UseArrowKeys    =   -1  'True
      BackGroundColor =   16761024
      BorderStyle     =   0
      Enabled         =   -1  'True
      Transparent     =   0   'False
      UpdateFrameRate =   60
      OptionSaveAllProperties=   0   'False
      AutoFrameRate   =   0   'False
      Object.Visible         =   -1  'True
      Object.Width           =   73
      Object.Height          =   89
      OPCItemCount    =   0
   End
   Begin iProfessionalLibrary.iSwitchRocker3WayX iSwitchRocker3WayX6 
      Height          =   1335
      Left            =   9000
      TabIndex        =   16
      Top             =   5520
      Width           =   1095
      ShowFocusRect   =   -1  'True
      BorderMargin    =   2
      RepeatInitialDelay=   500
      RepeatInterval  =   50
      Value           =   0
      Increment       =   1
      UseArrowKeys    =   -1  'True
      BackGroundColor =   16761024
      BorderStyle     =   0
      Enabled         =   -1  'True
      Transparent     =   0   'False
      UpdateFrameRate =   60
      OptionSaveAllProperties=   0   'False
      AutoFrameRate   =   0   'False
      Object.Visible         =   -1  'True
      Object.Width           =   73
      Object.Height          =   89
      OPCItemCount    =   0
   End
   Begin iProfessionalLibrary.iSwitchRocker3WayX iSwitchRocker3WayX7 
      Height          =   1335
      Left            =   9000
      TabIndex        =   17
      Top             =   7200
      Width           =   1095
      ShowFocusRect   =   -1  'True
      BorderMargin    =   2
      RepeatInitialDelay=   500
      RepeatInterval  =   50
      Value           =   0
      Increment       =   1
      UseArrowKeys    =   -1  'True
      BackGroundColor =   16761024
      BorderStyle     =   0
      Enabled         =   -1  'True
      Transparent     =   0   'False
      UpdateFrameRate =   60
      OptionSaveAllProperties=   0   'False
      AutoFrameRate   =   0   'False
      Object.Visible         =   -1  'True
      Object.Width           =   73
      Object.Height          =   89
      OPCItemCount    =   0
   End
   Begin iProfessionalLibrary.iSwitchRocker3WayX iSwitchRocker3WayX8 
      Height          =   1335
      Left            =   11040
      TabIndex        =   36
      Top             =   7200
      Width           =   1095
      ShowFocusRect   =   -1  'True
      BorderMargin    =   2
      RepeatInitialDelay=   500
      RepeatInterval  =   50
      Value           =   0
      Increment       =   10
      UseArrowKeys    =   -1  'True
      BackGroundColor =   16761024
      BorderStyle     =   0
      Enabled         =   -1  'True
      Transparent     =   0   'False
      UpdateFrameRate =   60
      OptionSaveAllProperties=   0   'False
      AutoFrameRate   =   0   'False
      Object.Visible         =   -1  'True
      Object.Width           =   73
      Object.Height          =   89
      OPCItemCount    =   0
   End
   Begin VB.Label Label25 
      Alignment       =   2  '가운데 맞춤
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
      ForeColor       =   &H00FFC0C0&
      Height          =   375
      Left            =   12120
      TabIndex        =   38
      Top             =   7200
      Width           =   615
   End
   Begin VB.Label Label24 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00200005&
      Caption         =   "Dn"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   375
      Left            =   12120
      TabIndex        =   37
      Top             =   8160
      Width           =   615
   End
   Begin VB.Label Label23 
      BackStyle       =   0  '투명
      Caption         =   "dBm"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   375
      Left            =   8160
      TabIndex        =   35
      Top             =   7560
      Width           =   735
   End
   Begin VB.Label Label22 
      BackStyle       =   0  '투명
      Caption         =   "MHz"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   375
      Left            =   8160
      TabIndex        =   34
      Top             =   4560
      Width           =   735
   End
   Begin VB.Label Label21 
      BackStyle       =   0  '투명
      Caption         =   "MHz"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   375
      Left            =   8160
      TabIndex        =   33
      Top             =   6120
      Width           =   735
   End
   Begin VB.Label Label20 
      BackStyle       =   0  '투명
      Caption         =   "MHz"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   375
      Left            =   8160
      TabIndex        =   32
      Top             =   3120
      Width           =   735
   End
   Begin VB.Image imgBack 
      Height          =   1950
      Left            =   13440
      Picture         =   "frmSetting.frx":5355
      Stretch         =   -1  'True
      Top             =   9600
      Width           =   1920
   End
   Begin VB.Label Label19 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00200005&
      Caption         =   "Dn"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   375
      Left            =   12120
      TabIndex        =   31
      Top             =   3480
      Width           =   615
   End
   Begin VB.Label Label18 
      Alignment       =   2  '가운데 맞춤
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
      ForeColor       =   &H00FFC0C0&
      Height          =   375
      Left            =   12120
      TabIndex        =   30
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label Label17 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00200005&
      Caption         =   "Dn"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   375
      Left            =   12120
      TabIndex        =   29
      Top             =   6480
      Width           =   615
   End
   Begin VB.Label Label16 
      Alignment       =   2  '가운데 맞춤
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
      ForeColor       =   &H00FFC0C0&
      Height          =   375
      Left            =   12120
      TabIndex        =   28
      Top             =   5520
      Width           =   615
   End
   Begin VB.Label Label15 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00200005&
      Caption         =   "Dn"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   375
      Left            =   12120
      TabIndex        =   27
      Top             =   4920
      Width           =   615
   End
   Begin VB.Label Label14 
      Alignment       =   2  '가운데 맞춤
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
      ForeColor       =   &H00FFC0C0&
      Height          =   375
      Left            =   12120
      TabIndex        =   26
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label Label13 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00200005&
      Caption         =   "Dn"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   375
      Left            =   10080
      TabIndex        =   25
      Top             =   8160
      Width           =   615
   End
   Begin VB.Label Label12 
      Alignment       =   2  '가운데 맞춤
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
      ForeColor       =   &H00FFC0C0&
      Height          =   375
      Left            =   10080
      TabIndex        =   24
      Top             =   7200
      Width           =   615
   End
   Begin VB.Label Label11 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00200005&
      Caption         =   "Dn"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   375
      Left            =   10080
      TabIndex        =   23
      Top             =   6480
      Width           =   615
   End
   Begin VB.Label Label10 
      Alignment       =   2  '가운데 맞춤
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
      ForeColor       =   &H00FFC0C0&
      Height          =   375
      Left            =   10080
      TabIndex        =   22
      Top             =   5520
      Width           =   615
   End
   Begin VB.Label Label9 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00200005&
      Caption         =   "Dn"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   375
      Left            =   10080
      TabIndex        =   21
      Top             =   4920
      Width           =   615
   End
   Begin VB.Label Label8 
      Alignment       =   2  '가운데 맞춤
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
      ForeColor       =   &H00FFC0C0&
      Height          =   375
      Left            =   10080
      TabIndex        =   20
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label Down 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00200005&
      Caption         =   "Dn"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   375
      Left            =   10080
      TabIndex        =   19
      Top             =   3480
      Width           =   615
   End
   Begin VB.Label Up 
      Alignment       =   2  '가운데 맞춤
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
      ForeColor       =   &H00FFC0C0&
      Height          =   375
      Left            =   10080
      TabIndex        =   18
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '투명
      Caption         =   "1/Step"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   375
      Left            =   9120
      TabIndex        =   10
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '투명
      Caption         =   "10/Step"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   375
      Left            =   11280
      TabIndex        =   9
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '투명
      Caption         =   "Amplitude (0~20dBm)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   1440
      TabIndex        =   4
      Top             =   7440
      Width           =   4455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '투명
      Caption         =   "Stop Frequency (2~80MHz)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   1440
      TabIndex        =   3
      Top             =   6000
      Width           =   5295
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '투명
      Caption         =   "Start Band (2~80MHz)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   4440
      Width           =   4215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '투명
      Caption         =   "Sampling Frequency"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   3000
      Width           =   4215
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "Spectrum Analyzer Setting"
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
      TabIndex        =   0
      Top             =   360
      Width           =   15135
   End
End
Attribute VB_Name = "frmSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer
Dim b As Integer
Dim c As Integer
Dim d As Integer
Dim e As Integer
Dim f As Integer
Dim g As Integer
Dim h As Integer
Dim Amp As Integer
Private Sub CmdBack_Click()
    frmMainForm.Show
    Unload Me
End Sub

Private Sub imgBack_Click()
    frmMainForm.Show
End Sub

Private Sub iSwitchRocker3WayX1_OnValueChange()
     f = iSwitchRocker3WayX1.Value
     Call cmd
End Sub

Private Sub iSwitchRocker3WayX2_OnValueChange()
    a = iSwitchRocker3WayX2.Value
    Call cmd
End Sub
Private Sub iSwitchRocker3WayX5_OnValueChange()
    b = iSwitchRocker3WayX5.Value
    Call cmd
End Sub
Private Sub iSwitchRocker3WayX3_OnValueChange()
    c = iSwitchRocker3WayX3.Value
    Call cmd
End Sub

Private Sub iSwitchRocker3WayX4_OnValueChange()
    e = iSwitchRocker3WayX4.Value
    Call cmd
End Sub
Private Sub cmd()
Dim StartBsum As Integer
Dim StopFsum As Integer

StartBsum = a + b
StopFsum = c + d
Amp = g + h

If StartBsum >= 0 And StartBsum <= 80 Then
    StartBtxt.Text = StartBsum
Else
    MsgBox "범위를 벗어납니다.", vbExclamation, "Setting"
    iSwitchRocker3WayX2.Value = 0
    iSwitchRocker3WayX5.Value = 0
    StartBsum = 0
    StartBtxt.Text = 0
End If
If StopFsum >= 0 And StopFsum <= 80 Then
    StopFtxt.Text = StopFsum
Else
    MsgBox "범위를 벗어납니다.", vbExclamation, "Setting"
    iSwitchRocker3WayX3.Value = 0
    iSwitchRocker3WayX6.Value = 0
    StopFsum = 0
    StopFtxt.Text = 0
End If
If StartBsum >= 0 And StartBsum <= 80 Then
    StartBtxt.Text = StartBsum
Else
    MsgBox "범위를 벗어납니다.", vbExclamation, "Setting"
    iSwitchRocker3WayX2.Value = 0
    iSwitchRocker3WayX5.Value = 0
    StartBsum = 0
    StartBtxt.Text = 0
End If
    SampFtxt.Text = e + f
If Amp >= 0 And Amp <= 20 Then
    Amptxt.Text = Amp
Else
    MsgBox "범위를 벗어납니다.", vbExclamation, "Setting"
    Amptxt.Text = 0
    iSwitchRocker3WayX7.Value = 0
    iSwitchRocker3WayX8.Value = 0
End If

End Sub

Private Sub iSwitchRocker3WayX6_OnValueChange()
    d = iSwitchRocker3WayX6.Value
    Call cmd
End Sub

Private Sub iSwitchRocker3WayX7_OnValueChange()
    g = iSwitchRocker3WayX7.Value
    Call cmd
End Sub

Private Sub iSwitchRocker3WayX8_OnValueChange()
    h = iSwitchRocker3WayX8.Value
    Call cmd
End Sub
