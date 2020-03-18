VERSION 5.00
Begin VB.Form frmReceive 
   BackColor       =   &H00200005&
   BorderStyle     =   0  '없음
   Caption         =   "Form1"
   ClientHeight    =   4005
   ClientLeft      =   1995
   ClientTop       =   4995
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txpowertxt 
      Alignment       =   2  '가운데 맞춤
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
      Left            =   3720
      TabIndex        =   0
      Text            =   "0 ˚"
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00200005&
      Caption         =   "Phase"
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
      Left            =   1200
      TabIndex        =   1
      Top             =   1440
      Width           =   2055
   End
End
Attribute VB_Name = "frmReceive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
