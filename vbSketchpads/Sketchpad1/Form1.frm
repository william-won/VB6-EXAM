VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picCanvas 
      Height          =   3135
      Left            =   0
      ScaleHeight     =   3075
      ScaleWidth      =   4635
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'A very simple sketchpad

Private Sub Form_Load()
    'Redraws the image  when the window is covered/uncovered
    picCanvas.AutoRedraw = True
    
    'The canvas look and drawing behaviour
    picCanvas.DrawWidth = 2 'Lines are 2 units thick
    picCanvas.ForeColor = vbBlue 'Lines are blue
    picCanvas.BackColor = vbWhite 'Background canvas is white
End Sub

'Set the first point and start the line
Private Sub picCanvas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        picCanvas.Line (X, Y)-(X, Y)
    End If
End Sub
'Continue the line to the next point as we move the mouse
Private Sub picCanvas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        picCanvas.Line -(X, Y)
    End If
End Sub
