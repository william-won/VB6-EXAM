VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   10680
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13155
   LinkTopic       =   "Form2"
   ScaleHeight     =   10680
   ScaleWidth      =   13155
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton cmdexit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   5640
      TabIndex        =   6
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton cmdlogin 
      Caption         =   "&Login"
      Height          =   495
      Left            =   3960
      TabIndex        =   5
      Top             =   5160
      Width           =   1215
   End
   Begin VB.TextBox txtpassword 
      Height          =   495
      IMEMode         =   3  '사용 못함
      Left            =   5640
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   4440
      Width           =   1215
   End
   Begin VB.TextBox txtuser 
      Height          =   495
      Left            =   5640
      TabIndex        =   3
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Password"
      Height          =   495
      Left            =   4080
      TabIndex        =   2
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "User Name"
      Height          =   495
      Left            =   4080
      TabIndex        =   1
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "LOGIN SCREEN"
      Height          =   975
      Left            =   4200
      TabIndex        =   0
      Top             =   2640
      Width           =   3855
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdlogin_Click()
If (txtuser = "HELLO" And txtpassword = "HELLO") Then
MsgBox "it is correct"
Form1.Show
Else
MsgBox "it is incorrect"
txtuser = ""
txtpassword = ""
txtuser.SetFocus
End If
End Sub


Private Sub txtpassword_LostFocus()
txtuser = UCase(txtuser)
txtpassword = UCase(txtpassword)
End Sub
