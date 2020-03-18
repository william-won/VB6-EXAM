VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "RFID를 이용한 출입통제 시스템"
   ClientHeight    =   5955
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   5265
   LinkTopic       =   "Form1"
   ScaleHeight     =   5955
   ScaleWidth      =   5265
   StartUpPosition =   3  'Windows 기본값
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   2400
      Top             =   240
   End
   Begin VB.Frame Frame3 
      Caption         =   "데이터 전송"
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
      Left            =   120
      TabIndex        =   5
      Top             =   3480
      Width           =   5055
      Begin VB.CommandButton Command3 
         Caption         =   "전송"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3960
         TabIndex        =   7
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00FFC0C0&
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
         Height          =   390
         Left            =   1440
         TabIndex        =   6
         Text            =   "2A4963"
         Top             =   600
         Width           =   2295
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H000000FF&
         FillStyle       =   0  '단색
         Height          =   375
         Left            =   3000
         Shape           =   3  '원형
         Top             =   1440
         Width           =   375
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H0000C000&
         FillStyle       =   0  '단색
         Height          =   375
         Left            =   1920
         Shape           =   3  '원형
         Top             =   1440
         Width           =   375
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H0000C0C0&
         FillStyle       =   0  '단색
         Height          =   375
         Left            =   840
         Shape           =   3  '원형
         Top             =   1440
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
         Left            =   2760
         TabIndex        =   12
         Top             =   1200
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
         Left            =   1680
         TabIndex        =   11
         Top             =   1200
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
         Left            =   600
         TabIndex        =   10
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label6 
         Alignment       =   2  '가운데 맞춤
         Caption         =   "카드검색"
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
         TabIndex        =   9
         Top             =   720
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
         Left            =   3960
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
      Begin VB.Shape STEP1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  '단색
         Height          =   375
         Left            =   4200
         Shape           =   3  '원형
         Top             =   600
         Width           =   375
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " 수신데이터 "
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
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   5055
      Begin VB.TextBox Text3 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00FFC0FF&
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
         TabIndex        =   15
         Top             =   1200
         Width           =   4815
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00FFC0FF&
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
         TabIndex        =   14
         Top             =   360
         Width           =   4815
      End
      Begin VB.Label Label3 
         Caption         =   "카드번호"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   960
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " 수신량 "
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
      Top             =   120
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
      Caption         =   "포트닫기"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "포트열기"
      Height          =   375
      Left            =   3960
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
   Begin MSCommLib.MSComm MSComm 
      Left            =   1680
      Top             =   240
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
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   5760
      Width           =   2535
   End
   Begin VB.Label Label10 
      Caption         =   "전자정보기술 4학년 원 대 희"
      Height          =   255
      Left            =   2760
      TabIndex        =   13
      Top             =   5760
      Width           =   2415
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00800000&
      FillColor       =   &H0080C0FF&
      FillStyle       =   0  '단색
      Height          =   495
      Left            =   3120
      Shape           =   3  '원형
      Top             =   360
      Width           =   735
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
         End
         Begin VB.Menu mnu_com2 
            Caption         =   "COM2"
            Checked         =   -1  'True
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
   Begin VB.Menu mnu_card 
      Caption         =   "카드"
      Begin VB.Menu mnu_creat 
         Caption         =   "등록"
      End
      Begin VB.Menu mnu_del 
         Caption         =   "삭제"
      End
      Begin VB.Menu mnu_reg 
         Caption         =   "등록된카드"
      End
   End
   Begin VB.Menu mnu_help 
      Caption         =   "도움말"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MSComm.PortOpen = True
Shape1.FillColor = &H80FF80
Command1.Enabled = False
Command2.Enabled = True
Command3.Enabled = True
Label2.Caption = "포트를 열었습니다.."
End Sub

Private Sub Command2_Click()
MSComm.PortOpen = False
Shape1.FillColor = &H80C0FF
Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = False
Label2.Caption = "포트를 닫았습니다.."
End Sub

Private Sub Command3_Click()
Text3.Text = ""
STEP1.FillColor = &HC0C0&
Text2.Text = ""
Label1.Caption = ""
Label2.Caption = "카드를 검색합니다.."
If ConvertTxt2Binary(Text1.Text, Form1) = True Then   'READ
End If
End Sub

Private Sub Command5_Click()
Timer1.Enabled = True
End Sub

Private Sub Form_Load()
MSComm.CommPort = 1
MSComm.Settings = "9600" + Setting
Return_Massage = 0
Timer1.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub


Private Sub mnu_creat_Click()
Reg_Card = InputBox("등록할 카드번호를 입력하십시오")
End Sub

Private Sub mnu_del_Click()
a = MsgBox("등록된 카드를 삭제 하시겠습니까?", vbYesNo)
If a = 6 Then
   Reg_Card = ""
End If
End Sub

Private Sub mnu_end_Click()
End
End Sub


Private Sub mnu_reg_Click()
MsgBox (Reg_Card)
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
                If Text3.Text = Reg_Card Then
                    Label2.Caption = "등록된 카드입니다"
                Else
                    Label2.Caption = "등록되지 않은 카드입니다"
                End If
            End If
End Sub

