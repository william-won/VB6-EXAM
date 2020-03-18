VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   8475
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16500
   LinkTopic       =   "Form4"
   ScaleHeight     =   8475
   ScaleWidth      =   16500
   StartUpPosition =   3  'Windows ±âº»°ª
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   360
      Top             =   2400
   End
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   360
      Picture         =   "Moving PicBox.frx":0000
      ScaleHeight     =   435
      ScaleWidth      =   2475
      TabIndex        =   0
      Top             =   4200
      Width           =   2535
      Begin VB.Label Label1 
         Caption         =   "HIT COMPUTER"
         Height          =   255
         Left            =   600
         TabIndex        =   1
         Top             =   120
         Width           =   1695
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer


Private Sub Form_Load()
a = 1
End Sub

Private Sub Timer1_Timer()
If a = 1 Then
    Picture1.Left = Picture1.Left + 100
    If Picture1.Left >= 8000 Then
    a = 2
    End If
ElseIf a = 2 Then
    Picture1.Left = Picture1.Left - 100
    If Picture1.Left <= 500 Then
    a = 3
    End If
End If
End Sub
