VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Form5"
   ClientHeight    =   9840
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13920
   LinkTopic       =   "Form5"
   ScaleHeight     =   9840
   ScaleWidth      =   13920
   StartUpPosition =   3  'Windows ±âº»°ª
   Begin VB.FileListBox File1 
      Height          =   4230
      Left            =   6360
      TabIndex        =   2
      Top             =   1440
      Width           =   2655
   End
   Begin VB.DirListBox Dir1 
      Height          =   4290
      Left            =   3480
      TabIndex        =   1
      Top             =   1440
      Width           =   2655
   End
   Begin VB.DriveListBox Drive1 
      Height          =   300
      Left            =   1560
      TabIndex        =   0
      Top             =   1440
      Width           =   1815
   End
   Begin VB.OLE OLE1 
      Class           =   "Excel.Sheet.12"
      Height          =   3015
      Left            =   2880
      OleObjectBlob   =   "DriverFile.frx":0000
      SourceDoc       =   "C:\123.xlsx"
      TabIndex        =   3
      Top             =   6000
      Width           =   6735
   End
   Begin VB.Image Image1 
      Height          =   2895
      Left            =   3480
      Stretch         =   -1  'True
      Top             =   6000
      Width           =   5535
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
'Image1.Picture = LoadPicture(Dir1.Path & "\" & File1.FileName)
OLE1.CreateLink Dir1.Path & "\" & File1.FileName
End Sub
