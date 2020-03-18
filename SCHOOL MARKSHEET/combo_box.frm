VERSION 5.00
Begin VB.Form Form7 
   Caption         =   "Form7"
   ClientHeight    =   8565
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13395
   LinkTopic       =   "Form7"
   ScaleHeight     =   8565
   ScaleWidth      =   13395
   StartUpPosition =   3  'Windows ±âº»°ª
   Begin VB.ComboBox Combo2 
      Height          =   300
      Left            =   3960
      TabIndex        =   14
      Text            =   "Combo2"
      Top             =   5040
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "combo_box.frx":0000
      Left            =   3960
      List            =   "combo_box.frx":000D
      Sorted          =   -1  'True
      TabIndex        =   12
      Text            =   "Combo1"
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "account type"
      Height          =   975
      Left            =   3960
      TabIndex        =   9
      Top             =   3360
      Width           =   3255
      Begin VB.OptionButton Option4 
         Caption         =   "current"
         Height          =   495
         Left            =   1800
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Option3 
         Caption         =   "saving"
         Height          =   495
         Left            =   480
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "gender"
      Height          =   855
      Left            =   3960
      TabIndex        =   6
      Top             =   2160
      Width           =   3255
      Begin VB.OptionButton Option1 
         Caption         =   "male"
         Height          =   495
         Left            =   480
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "female"
         Height          =   495
         Left            =   1800
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Korean"
      Height          =   495
      Left            =   6720
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Hindi"
      Height          =   495
      Left            =   5400
      TabIndex        =   1
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "English"
      Height          =   495
      Left            =   4080
      TabIndex        =   0
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Country"
      Height          =   495
      Left            =   2760
      TabIndex        =   13
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "account type"
      Height          =   495
      Left            =   2640
      TabIndex        =   5
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "gender"
      Height          =   495
      Left            =   2760
      TabIndex        =   4
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Lanfuages Known"
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo2_GotFocus()
Combo2.Clear
Combo2.AddItem "A+"
Combo2.AddItem "A-"
Combo2.AddItem "O+"
Combo2.AddItem "O-"
End Sub
