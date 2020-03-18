VERSION 5.00
Begin VB.Form Form8 
   Caption         =   "Form8"
   ClientHeight    =   12915
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   20745
   LinkTopic       =   "Form8"
   ScaleHeight     =   12915
   ScaleWidth      =   20745
   StartUpPosition =   3  'Windows 기본값
   Begin VB.Frame Frame2 
      Caption         =   "Font Color"
      Height          =   1215
      Left            =   11760
      TabIndex        =   15
      Top             =   1080
      Visible         =   0   'False
      Width           =   3015
      Begin VB.CommandButton Command11 
         BackColor       =   &H00000000&
         Height          =   350
         Left            =   600
         Style           =   1  '그래픽
         TabIndex        =   18
         Top             =   600
         Width           =   500
      End
      Begin VB.CommandButton Command10 
         BackColor       =   &H00FFFFFF&
         Height          =   350
         Left            =   1200
         Style           =   1  '그래픽
         TabIndex        =   17
         Top             =   600
         Width           =   500
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H0000FFFF&
         Height          =   350
         Left            =   1800
         Style           =   1  '그래픽
         TabIndex        =   16
         Top             =   600
         Width           =   500
      End
   End
   Begin VB.CommandButton Command8 
      Height          =   495
      Left            =   11760
      Picture         =   "Font_Size_Color.frx":0000
      Style           =   1  '그래픽
      TabIndex        =   14
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton Command7 
      Height          =   495
      Left            =   8640
      Picture         =   "Font_Size_Color.frx":0152
      Style           =   1  '그래픽
      TabIndex        =   13
      Top             =   480
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Caption         =   "Back Color"
      Height          =   1215
      Left            =   8640
      TabIndex        =   9
      Top             =   1080
      Visible         =   0   'False
      Width           =   3015
      Begin VB.CommandButton Command6 
         BackColor       =   &H00FF00FF&
         Height          =   350
         Left            =   1800
         Style           =   1  '그래픽
         TabIndex        =   12
         Top             =   600
         Width           =   500
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H0000C000&
         Height          =   350
         Left            =   1200
         Style           =   1  '그래픽
         TabIndex        =   11
         Top             =   600
         Width           =   500
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H000000FF&
         Height          =   350
         Left            =   600
         Style           =   1  '그래픽
         TabIndex        =   10
         Top             =   600
         Width           =   500
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "R"
      Height          =   495
      Left            =   7800
      TabIndex        =   8
      Top             =   720
      Width           =   500
   End
   Begin VB.CommandButton Command2 
      Caption         =   "C"
      Height          =   495
      Left            =   7200
      TabIndex        =   7
      Top             =   720
      Width           =   500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "L"
      Height          =   495
      Left            =   6600
      TabIndex        =   6
      Top             =   720
      Width           =   500
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      ItemData        =   "Font_Size_Color.frx":02A4
      Left            =   5280
      List            =   "Font_Size_Color.frx":02B4
      TabIndex        =   5
      Text            =   "Combo2"
      Top             =   840
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "Font_Size_Color.frx":02F0
      Left            =   3960
      List            =   "Font_Size_Color.frx":02F2
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   840
      Width           =   1215
   End
   Begin VB.CheckBox Check3 
      Caption         =   "U"
      Height          =   495
      Left            =   2880
      TabIndex        =   3
      Top             =   720
      Width           =   800
   End
   Begin VB.CheckBox Check2 
      Caption         =   "I"
      Height          =   495
      Left            =   1920
      TabIndex        =   2
      Top             =   720
      Width           =   800
   End
   Begin VB.CheckBox Check1 
      Caption         =   "B"
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   720
      Width           =   800
   End
   Begin VB.TextBox Text1 
      Height          =   10335
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Font_Size_Color.frx":02F4
      Top             =   2520
      Width           =   20535
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.Value = 1 Then
Text1.FontBold = True
Else
Text1.FontBold = False
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
Text1.FontItalic = True
Else
Text1.FontItalic = False
End If
End Sub

Private Sub Check3_Click()
If Check3.Value = 1 Then
Text1.FontUnderline = True
Else
Text1.FontUnderline = False
End If
End Sub

Private Sub Combo1_Click()
Text1.FontSize = Combo1
End Sub

Private Sub Combo1_GotFocus()
Combo1.Clear
For a = 2 To 72
Combo1.AddItem a
Next
End Sub

Private Sub Combo2_Click()
Text1.FontName = Combo2
End Sub

Private Sub Command1_Click()
Text1.Alignment = 0
End Sub

Private Sub Command10_Click()
Text1.ForeColor = vbWhite
Frame2.Visible = False
End Sub

Private Sub Command11_Click()
Text1.ForeColor = vbBlack
Frame2.Visible = False
End Sub

Private Sub Command2_Click()
Text1.Alignment = 2
End Sub

Private Sub Command3_Click()
Text1.Alignment = 1
End Sub


Private Sub Command4_Click()
Text1.BackColor = vbRed
Frame1.Visible = False
End Sub

Private Sub Command5_Click()
Text1.BackColor = vbGreen
Frame1.Visible = False
End Sub

Private Sub Command6_Click()
Text1.BackColor = &HFF00FF
Frame1.Visible = False
End Sub

Private Sub Command7_Click()
Frame1.Visible = True
End Sub

Private Sub Command8_Click()
Frame2.Visible = True
End Sub

Private Sub Command9_Click()
Text1.ForeColor = vbYellow
Frame2.Visible = False
End Sub
