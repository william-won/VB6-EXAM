VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "RFID를 이용한 신원조회 시스템"
   ClientHeight    =   10035
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   12015
   LinkTopic       =   "Form1"
   ScaleHeight     =   10035
   ScaleWidth      =   12015
   StartUpPosition =   3  'Windows 기본값
   Begin VB.Frame Frame4 
      Caption         =   "신원조회"
      ForeColor       =   &H00008000&
      Height          =   6135
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   11775
      Begin VB.Frame Frame6 
         Caption         =   "교통기록"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   2535
         Left            =   6000
         TabIndex        =   38
         Top             =   3360
         Width           =   5535
         Begin VB.TextBox Text16 
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1920
            TabIndex        =   48
            Top             =   1920
            Width           =   3375
         End
         Begin VB.TextBox Text15 
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1920
            TabIndex        =   47
            Top             =   1200
            Width           =   3375
         End
         Begin VB.TextBox Text14 
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1920
            TabIndex        =   46
            Top             =   480
            Width           =   3375
         End
         Begin VB.Label Label23 
            Caption         =   "벌     점 :"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   11.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   45
            Top             =   1920
            Width           =   1215
         End
         Begin VB.Label Label22 
            Caption         =   "삼진아웃 :"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   11.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   44
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label Label21 
            Caption         =   "사고경력 :"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   11.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   43
            Top             =   480
            Width           =   1215
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "전과기록"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   2535
         Left            =   240
         TabIndex        =   35
         Top             =   3360
         Width           =   5415
         Begin VB.TextBox Text13 
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1920
            TabIndex        =   42
            Top             =   1920
            Width           =   2895
         End
         Begin VB.TextBox Text12 
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1920
            TabIndex        =   41
            Top             =   1200
            Width           =   2895
         End
         Begin VB.TextBox Text11 
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1920
            TabIndex        =   40
            Top             =   480
            Width           =   2895
         End
         Begin VB.Label Label20 
            Caption         =   "기타사항 :"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   11.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   39
            Top             =   1920
            Width           =   1215
         End
         Begin VB.Label Label19 
            Caption         =   "지명수배 :"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   11.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   37
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label Label18 
            Caption         =   "범죄경력 :"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   11.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   36
            Top             =   480
            Width           =   1215
         End
      End
      Begin VB.TextBox Text10 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8400
         TabIndex        =   34
         Top             =   2760
         Width           =   3135
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         TabIndex        =   32
         Top             =   2760
         Width           =   3375
      End
      Begin VB.TextBox Text9 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8400
         TabIndex        =   30
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4080
         TabIndex        =   28
         Top             =   2040
         Width           =   2655
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         TabIndex        =   26
         Top             =   1320
         Width           =   7935
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8760
         TabIndex        =   25
         Top             =   600
         Width           =   2775
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6360
         TabIndex        =   21
         Top             =   600
         Width           =   1215
      End
      Begin VB.PictureBox Picture1 
         Height          =   2320
         Left            =   240
         Picture         =   "Form1.frx":0000
         ScaleHeight     =   2265
         ScaleWidth      =   1710
         TabIndex        =   18
         Top             =   360
         Width           =   1770
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4080
         TabIndex        =   17
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label17 
         Caption         =   "핸드폰 :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7320
         TabIndex        =   33
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "E-Mail :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   31
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label16 
         Caption         =   "혈액형 :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7320
         TabIndex        =   29
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label15 
         Caption         =   "주민등록번호 :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   27
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label14 
         Caption         =   "(영문)"
         Height          =   255
         Left            =   8040
         TabIndex        =   24
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label13 
         Caption         =   "(한문)"
         Height          =   255
         Left            =   5640
         TabIndex        =   23
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label12 
         Caption         =   "(한글)"
         Height          =   255
         Left            =   3360
         TabIndex        =   22
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "현 주 소 :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   20
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "성  명 :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   19
         Top             =   600
         Width           =   855
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   6840
      Top             =   9480
   End
   Begin VB.Frame Frame3 
      Caption         =   "Data Transmit"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   2175
      Left            =   1560
      TabIndex        =   5
      Top             =   7200
      Width           =   5055
      Begin VB.CommandButton Command3 
         Caption         =   "CARD Reading 대기모드로 전환 (&A)"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   6
         Top             =   1440
         Width           =   4215
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H000000FF&
         FillStyle       =   0  '단색
         Height          =   375
         Left            =   4200
         Shape           =   3  '원형
         Top             =   480
         Width           =   375
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H0000C000&
         FillStyle       =   0  '단색
         Height          =   375
         Left            =   2880
         Shape           =   3  '원형
         Top             =   480
         Width           =   375
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H0000C0C0&
         FillStyle       =   0  '단색
         Height          =   375
         Left            =   1680
         Shape           =   3  '원형
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label9 
         Alignment       =   2  '가운데 맞춤
         Caption         =   "ERROR"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   10
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label8 
         Alignment       =   2  '가운데 맞춤
         Caption         =   "RX"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   9
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label7 
         Alignment       =   2  '가운데 맞춤
         Caption         =   "TX"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   8
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label5 
         Alignment       =   2  '가운데 맞춤
         Caption         =   "수신여부"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   855
      End
      Begin VB.Shape STEP1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  '단색
         Height          =   375
         Left            =   480
         Shape           =   3  '원형
         Top             =   480
         Width           =   375
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "CARD No."
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2175
      Left            =   6840
      TabIndex        =   4
      Top             =   7200
      Width           =   5055
      Begin VB.TextBox Text3 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   36
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   4815
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   1680
         Width           =   4815
      End
      Begin VB.Label Label3 
         Caption         =   "Receive Data"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1320
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "RX Bit"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   7200
      Width           =   1335
      Begin VB.Label Label1 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "시스템해제(&Z)"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   9000
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "시스템연결(&Q)"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   8520
      Width           =   1335
   End
   Begin MSCommLib.MSComm MSComm 
      Left            =   6240
      Top             =   9480
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      RThreshold      =   1
      InputMode       =   1
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   20.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   495
      Left            =   240
      TabIndex        =   14
      Top             =   6480
      Width           =   11535
   End
   Begin VB.Label Label10 
      Alignment       =   1  '오른쪽 맞춤
      Caption         =   "전자정보기술공학과  4학년  원 대 희"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6960
      TabIndex        =   11
      Top             =   9600
      Width           =   4935
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  '단색
      Height          =   255
      Left            =   120
      Top             =   8160
      Width           =   1335
   End
   Begin VB.Menu mnu_file 
      Caption         =   "파일"
      Begin VB.Menu mnu_end 
         Caption         =   "종료"
      End
   End
   Begin VB.Menu mnu_set 
      Caption         =   "설정"
      Begin VB.Menu mnu_port 
         Caption         =   "포트설정"
         Begin VB.Menu mnu_com1 
            Caption         =   "COM1"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnu_com2 
            Caption         =   "COM2"
         End
      End
      Begin VB.Menu mnu_baud 
         Caption         =   "속도설정"
         Begin VB.Menu mnu_9600 
            Caption         =   "9600bps"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnu_115200 
            Caption         =   "115200bps"
         End
      End
   End
   Begin VB.Menu mnu_made 
      Caption         =   "제작후기"
      Begin VB.Menu mnu_push 
         Caption         =   "꾸욱! 눌러주세요~~"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()    '시스템 연결버튼에 대한 함수
MSComm.PortOpen = True          '포트를 연다.

Shape1.FillColor = &HFF00&      'Shape1의 색깔을 파란색으로 바꾼다(전원ON)
Command1.Enabled = False        '시스템연결 버튼을 비활성화
Command2.Enabled = True         '시스템해제 버튼을 활성화
Command3.Enabled = True         '카드읽기 버튼을 활성화
Label2.Caption = "시스템과 연결이 되었습니다."      '화면 아래 부분에 연결이 되었음을 표시한다.

Picture1.Picture = LoadPicture("normal.jpg", 4, 0, 114, 152)    '사진 초기화
Text1.Text = ""                 'Text1 초기화
Text2.Text = ""                 'Text2 초기화
Text3.Text = ""                 'Text3 초기화
Text4.Text = ""                 'Text4 초기화
Text5.Text = ""                 'Text5 초기화
Text6.Text = ""                 'Text6 초기화
Text7.Text = ""                 'Text7 초기화
Text8.Text = ""                 'Text8 초기화
Text9.Text = ""                 'Text9 초기화
Text10.Text = ""                'Text10 초기화
Text11.Text = ""                'Text11 초기화
Text12.Text = ""                'Text12 초기화
Text13.Text = ""                'Text13 초기화
Text14.Text = ""                'Text14 초기화
Text15.Text = ""                'Text15 초기화
Text16.Text = ""                'Text16 초기화
Label1.Caption = ""             'RX Bit 초기화


End Sub

Private Sub Command2_Click()    '시스템 해제버튼에 대한 함수

MSComm.PortOpen = False         '포트를 닫는다.
Shape1.FillColor = &HFF&        '빨간색으로 바꾼다.(전원OFF)
STEP1.FillColor = &HFF&         '수신여부 LED소등
Command1.Enabled = True         '시스템연결 버튼을 활성화
Command2.Enabled = False        '시스템해제 버튼을 비활성화
Command3.Enabled = False        '카드읽기 버튼을 비활성화
Label2.Caption = "시스템과 연결이 해제되었습니다."      '화면 아래 부분에 해제 되었음을 표시한다.

Picture1.Picture = LoadPicture("normal.jpg", 4, 0, 114, 152)    '사진 초기화
Text1.Text = ""                 'Text1 초기화
Text2.Text = ""                 'Text2 초기화
Text3.Text = ""                 'Text3 초기화
Text4.Text = ""                 'Text4 초기화
Text5.Text = ""                 'Text5 초기화
Text6.Text = ""                 'Text6 초기화
Text7.Text = ""                 'Text7 초기화
Text8.Text = ""                 'Text8 초기화
Text9.Text = ""                 'Text9 초기화
Text10.Text = ""                'Text10 초기화
Text11.Text = ""                'Text11 초기화
Text12.Text = ""                'Text12 초기화
Text13.Text = ""                'Text13 초기화
Text14.Text = ""                'Text14 초기화
Text15.Text = ""                'Text15 초기화
Text16.Text = ""                'Text16 초기화
Label1.Caption = ""             'RX Bit 초기화

End Sub

Private Sub Command3_Click()    '데이터 전송에 관한 함수
STEP1.FillColor = &HFF00&       '수신여부 LED점등

Picture1.Picture = LoadPicture("normal.jpg", 4, 0, 114, 152)    '사진 초기화
Text1.Text = ""                 'Text1 초기화
Text2.Text = ""                 'Text2 초기화
Text3.Text = ""                 'Text3 초기화
Text4.Text = ""                 'Text4 초기화
Text5.Text = ""                 'Text5 초기화
Text6.Text = ""                 'Text6 초기화
Text7.Text = ""                 'Text7 초기화
Text8.Text = ""                 'Text8 초기화
Text9.Text = ""                 'Text9 초기화
Text10.Text = ""                'Text10 초기화
Text11.Text = ""                'Text11 초기화
Text12.Text = ""                'Text12 초기화
Text13.Text = ""                'Text13 초기화
Text14.Text = ""                'Text14 초기화
Text15.Text = ""                'Text15 초기화
Text16.Text = ""                'Text16 초기화
Label1.Caption = ""             'RX Bit 초기화

Label2.Caption = "카드를 리더기에 대어주세요."
If ConvertTxt2Binary("2A4963", Form1) = True Then   '읽기 프로토콜 전송
End If
End Sub

Private Sub Form_Load()
MSComm.CommPort = 11
MSComm.Settings = "9600" + Setting
Return_Massage = 0
Timer1.Enabled = False
End Sub


Private Sub mnu_push_Click()
    Dim msg, button, title
    msg = "안녕하세요~ ^^*" & Chr(10) & Chr(10) & "이 작품을 만드느라 힘들었지만 재미있었습니다." & Chr(10) & Chr(10) & "기억속에 오래동안 남을 것입니다." & Chr(10) & Chr(10) & "감사합니다./(^.^)(_ _)(^.^)/"
    button = vbOKOnly
    title = "제 작 후 기"
    
    MsgBox msg, button, title
    
End Sub

Private Sub MSComm_OnComm()
    Select Case MSComm.CommEvent
         Case comEvReceive
            Dim buffer() As Byte
            Dim count As Integer
            Dim TempStr As String
            Dim RX_Count As Integer
            TempStr = ""
            
            buffer = MSComm.Input
            
            For count = 0 To UBound(buffer)
                tempchar = Hex(buffer(count))
                If Len(tempchar) = 1 Then tempchar = "0" + tempchar
                TempStr = TempStr + tempchar
            Next count
        
            Text2.Text = Text2.Text + TempStr
            STEP1.FillColor = &HC000&
            RX_Count = RX_Count + 1
            Text3.Text = Mid(Text2.Text, 5, 10)
            Label1.Caption = Len(Text2.Text) / 2
    End Select
                            
            If Len(Text2.Text) / 2 = 8 Then
            
                If Text3.Text = "0F0059373E" Then
                    Picture1.Picture = LoadPicture("oneday798.jpg", 4, 0, 114, 152)
                    Text1.Text = "oneday798"
                    Text4.Text = "원대희"
                    Text5.Text = "元大喜"
                    Text6.Text = "Won Dae Hui"
                    Text7.Text = "경기도 용인시 처인구 삼가동 늘푸른오스카빌 104-1401호"
                    Text8.Text = "790827-1xxxxxx"
                    Text9.Text = "AB형"
                    Text10.Text = "010-3846-5942"
                    Text11.Text = "전과가 없음"
                    Text12.Text = "대상이 아님"
                    Text13.Text = "없음"
                    Text14.Text = "1회(고속도로3중추돌)"
                    Text15.Text = "1아웃(알콜농도 0.045%)"
                    Text16.Text = "70점"
                    Label2.Caption = "원대희 님의 정보가 조회되었습니다."
                    End If
                
                If Text3.Text = "0F00595CA2" Then
                    Picture1.Picture = LoadPicture("ljh.jpg", 4, 0, 114, 152)
                    Text1.Text = "unc79@hanmail.net"
                    Text4.Text = "이준혁"
                    Text5.Text = "李準赫"
                    Text6.Text = "Lee Jun Hyuk"
                    Text7.Text = "서울시 방화3동 치현마을 서광A 105동 303호"
                    Text8.Text = "790830-1xxxxxx"
                    Text9.Text = "A형"
                    Text10.Text = "011-9941-6805"
                    Text11.Text = "전과가 없음"
                    Text12.Text = "대상이 아님"
                    Text13.Text = "없음"
                    Text14.Text = "없음"
                    Text15.Text = "2아웃(알콜농도 0.04%,0.05%)"
                    Text16.Text = "90점"
                    Label2.Caption = "이준혁 님의 정보가 조회되었습니다."
                    End If

         End If

 End Sub

