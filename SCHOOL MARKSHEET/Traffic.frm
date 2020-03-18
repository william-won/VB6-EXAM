VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   9240
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15285
   LinkTopic       =   "Form3"
   ScaleHeight     =   9240
   ScaleWidth      =   15285
   StartUpPosition =   3  'Windows ±âº»°ª
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   2280
      Top             =   3600
   End
   Begin VB.Image Image3 
      Height          =   1995
      Left            =   5000
      Picture         =   "Traffic.frx":0000
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   1995
   End
   Begin VB.Image Image2 
      Height          =   1995
      Left            =   5000
      Picture         =   "Traffic.frx":0442
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   1995
   End
   Begin VB.Image Image1 
      Height          =   1995
      Left            =   5000
      Picture         =   "Traffic.frx":0884
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   1995
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
If Image1.Visible = True Then
    Image1.Visible = False
    Image2.Visible = True
    Image3.Visible = False
ElseIf Image2.Visible = True Then
    Image1.Visible = False
    Image2.Visible = False
    Image3.Visible = True
Else
    Image1.Visible = True
    Image2.Visible = False
    Image3.Visible = False
End If
End Sub
