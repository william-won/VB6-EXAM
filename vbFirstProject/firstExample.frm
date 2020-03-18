VERSION 5.00
Begin VB.Form frmFirstExampleForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "My First Example"
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5475
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton ctrlPressMeButton 
      Caption         =   "Press Me"
      Height          =   615
      Left            =   180
      TabIndex        =   1
      ToolTipText     =   "Go ahead press me...I know you want to."
      Top             =   120
      Width           =   5115
   End
   Begin VB.CheckBox ctrlCheckBox1 
      Caption         =   "Red Form"
      Height          =   255
      Left            =   180
      TabIndex        =   0
      ToolTipText     =   "Check me to change the backgroun"
      Top             =   1440
      Width           =   2115
   End
   Begin VB.Label ctrlCoordinateLabel 
      Caption         =   "Mouse coordinates:"
      Height          =   375
      Left            =   180
      TabIndex        =   2
      ToolTipText     =   "Displays the current x, y coordinates (doesn't work over controls though)"
      Top             =   840
      Width           =   5115
   End
End
Attribute VB_Name = "frmFirstExampleForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Times As Integer 'The number of times the button is clicked

' This method is raised wehn the checkbox is clicked.
' If the check box is checked then the background color
' of the form will be set to red.  If the check box is not checked then
' the background color of the form will be set back to grey.
Private Sub ctrlCheckBox1_Click()

    ' If CheckBox checked then set background of form to red.
    If (ctrlCheckBox1.Value = Checked) Then
        frmFirstExampleForm.BackColor = RGB(255, 0, 0)
        
    ' If CheckBox is unchecked then set background of form to grey.
    ElseIf (ctrlCheckBox1.Value = Unchecked) Then
            frmFirstExampleForm.BackColor = RGB(210, 210, 210)
    End If
End Sub


' This method will change the text displayed in the button
' when the button is pressed.
Private Sub ctrlPressMeButton_Click()
    Times = Times + 1
    ctrlPressMeButton.Caption = "You pressed me --- click #" & Times
End Sub


' This method is executed whenever the Form detects a mouse move.
' The text box will display the current x,y coordinates of
' the mouse pointer even time that the mouse moves (generating a mouse move
' event). It won't update the mouse coordinates when the pointer moves over a
' VB control (such as a button though).  The current x,y coordinates for the
' mouse are automatically passed to this method.
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ctrlCoordinateLabel.Caption = "Mouse coordinates: (" & X & "," & Y & ")"
End Sub


