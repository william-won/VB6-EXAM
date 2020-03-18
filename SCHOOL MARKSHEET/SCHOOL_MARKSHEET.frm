VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5850
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5820
   LinkTopic       =   "Form1"
   ScaleHeight     =   5850
   ScaleWidth      =   5820
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton Command1 
      Caption         =   "Calculation"
      Height          =   495
      Left            =   2760
      TabIndex        =   15
      Top             =   5160
      Width           =   1215
   End
   Begin VB.TextBox txtavg 
      Height          =   495
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   4560
      Width           =   1215
   End
   Begin VB.TextBox txttotal 
      Height          =   495
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox txtmath 
      Height          =   495
      Left            =   2760
      TabIndex        =   12
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox txtkor 
      Height          =   495
      Left            =   2760
      TabIndex        =   11
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox txteng 
      Height          =   495
      Left            =   2760
      TabIndex        =   10
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2760
      TabIndex        =   9
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2760
      MaxLength       =   3
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "Average"
      Height          =   495
      Left            =   1320
      TabIndex        =   8
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Total"
      Height          =   495
      Left            =   1320
      TabIndex        =   7
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Maths"
      Height          =   495
      Left            =   1320
      TabIndex        =   6
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Korean"
      Height          =   495
      Left            =   1320
      TabIndex        =   5
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "English"
      Height          =   495
      Left            =   1320
      TabIndex        =   4
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Name"
      Height          =   495
      Left            =   1320
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Roll No"
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  '단일 고정
      Caption         =   "SCHOOL MARKSHEET"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   21.75
         Charset         =   129
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4845
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
txttotal = Val(txteng) + Val(txtkor) + Val(txtmath)
txtavg = txttotal / 3
End Sub

Private Sub txtmath_LostFocus()
txttotal = Val(txteng) + Val(txtkor) + Val(txtmath)
txtavg = txttotal / 3
End Sub
