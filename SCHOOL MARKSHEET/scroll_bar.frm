VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "Form6"
   ClientHeight    =   9480
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16200
   LinkTopic       =   "Form6"
   ScaleHeight     =   9480
   ScaleWidth      =   16200
   StartUpPosition =   3  'Windows ±âº»°ª
   Begin VB.VScrollBar VScroll1 
      Height          =   1215
      LargeChange     =   2000
      Left            =   6480
      Max             =   20000
      Min             =   10000
      SmallChange     =   1000
      TabIndex        =   3
      Top             =   2520
      Value           =   10000
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   1000
      Left            =   4920
      Max             =   5000
      Min             =   1000
      SmallChange     =   500
      TabIndex        =   2
      Top             =   3360
      Value           =   1000
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   4920
      TabIndex        =   0
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "salary"
      Height          =   495
      Left            =   3720
      TabIndex        =   1
      Top             =   2760
      Width           =   1215
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub HScroll1_Change()
Text1 = HScroll1.Value
End Sub

Private Sub VScroll1_Change()
Text1 = VScroll1.Value
End Sub
