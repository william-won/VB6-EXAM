VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Sketchpad Version 2"
   ClientHeight    =   4530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6045
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   ScaleHeight     =   4530
   ScaleWidth      =   6045
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optLineWidth 
      Caption         =   "Option1"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   4
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   4080
      Width           =   975
   End
   Begin VB.PictureBox picCanvas 
      Height          =   4335
      Left            =   1200
      ScaleHeight     =   4275
      ScaleWidth      =   4635
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
   Begin VB.Label lblLineWidth 
      BackColor       =   &H00000000&
      Height          =   135
      Index           =   0
      Left            =   360
      TabIndex        =   3
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label lblColor 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This sketchpad gives the user the ability to change line color and thickness

'The program illustrates several 'advanced' VB graphics
'First, controls are dynamically created through the load command; thus
' the programmer can easily add or decrease the number of colors and line
' thicknesses simply by altering the array that holds these values
'Second, the dynamic controls are programmatically positioned so they are in
' the correct location
'Third, the label control is used to mimic a color chip by how we show/hide its border
'Fourth, resizing the window adjusts the size of the canvas to fill the space

'An explanation about creating controls at run time
'The easiest way to do this is to create a control array
' i.e., add a control to a form and set its index property to 0.
'Doing a Load <control> will create a new control whose properties are somewhat similar
'to the original one; it will automatically have a new index that is 1 higher than the previous.
'For example, if you have a button named Command1 and set its index to 0,
'then 'Load Command1' will create a new button with its index set to 1. You can refere
'to these controls like an array i.e.,
' Command1(0).Caption = "First Button"
' Command1(1).Caption = "2nd Button"

Option Explicit

Const Margin As Integer = 10        'The margin between the left/right/top/bottom edges of the display
Const SmallSpace As Integer = 2     'Constants we use to space things apart vertically on the form
Const LargeSpace As Integer = 10

'Array containing the colors in the color pallette
Const MaxColors As Integer = 4      'The number of colors we will display + 1
Dim Colors(MaxColors + 1) As Long

'Array containing the widths in the line width pallette
Const MaxLineWidths As Integer = 4      'The number of line widths we will display + 1
Dim LineWidths(MaxLineWidths + 1) As Integer


'Initialize various paramaters and create / position the color and linewidth controls
Private Sub Form_Load()
    Dim I As Integer
    
    picCanvas.AutoRedraw = True   'Redraws the image  when the window is covered/uncovered
    picCanvas.BackColor = vbWhite 'Background canvas is white
    picCanvas.ToolTipText = "Click the left mouse button to sketch"
    cmdClear.ToolTipText = "Click to clear the canvas"
    Form1.ScaleMode = vbPixels    'Measures are all in pixels
    
    CreateColorPalette            'Create the Color and LineWidth Palette on the Form
    CreateLineWidthPalette
    
    'Set the initial selections of the Palettes
    lblColor_Click MaxColors      'Lines are black
    optLineWidth(0).Value = True  'Line thickness default is that of the first one shown
    optLineWidth_Click 0
End Sub

'Whenever the form is resized, make sure the canvas fills the empty space.
'We should also enforce a minimum window size, but we don't
Private Sub Form_Resize()
    picCanvas.Width = Form1.ScaleWidth - picCanvas.Left - Margin
    picCanvas.Top = Margin
    picCanvas.Height = Form1.ScaleHeight - picCanvas.Top - Margin
    cmdClear.Top = Form1.ScaleHeight - cmdClear.Height - Margin
End Sub


'''''
''''' Routines for creating the palette of controls
'''''

'Create the Color Palette
Private Sub CreateColorPalette()
    Dim I As Integer
    
    Colors(0) = vbYellow     'We first define our colors for the palette
    Colors(1) = vbGreen
    Colors(2) = vbBlue
    Colors(3) = vbRed
    Colors(4) = vbBlack
    
    lblColor(0).Left = Margin      'Align the first color label to the top left margin corner
    lblColor(0).Top = Margin       'All other controls will be positioned relative to this
    lblColor(0).ToolTipText = "Click to set the drawing color"
    lblColor(0).BackColor = Colors(0)   'We then set the first label to the first color, and
    For I = 1 To MaxColors              'then we create new labels corresponding to the other colors.
        Load lblColor(I)
        lblColor(I).BackColor = Colors(I) 'Display the color and reposition the label
        lblColor(I).Left = Margin
        lblColor(I).Top = lblColor(I - 1).Top + lblColor(I - 1).Height
        lblColor(I).Visible = True        'by default, load sets control visibility to False
    Next I
End Sub
'Create the LineWidth Pallette
'This works similar to CreateColor Palette, except we use two controls for this:
' an option control for showing which width is selected, and a label for displaying the width
Private Sub CreateLineWidthPalette()
    Dim I As Integer
    
    LineWidths(0) = 1   'Define the line widths for the palette
    LineWidths(1) = 2
    LineWidths(2) = 4
    LineWidths(3) = 7
    LineWidths(4) = 10
    
    lblLineWidth(0).ToolTipText = "Click to select a line width"
    optLineWidth(0).ToolTipText = lblLineWidth(0).ToolTipText
    
    For I = 0 To MaxLineWidths
        If I > 0 Then
            Load optLineWidth(I)    'If its not the first control, then create and position it
            Load lblLineWidth(I)
            optLineWidth(I).Top = optLineWidth(I - 1).Top + optLineWidth(I - 1).Height + SmallSpace
        Else
            optLineWidth(I).Top = lblColor(MaxColors).Top + lblColor(MaxColors).Height + LargeSpace
        End If
        optLineWidth(I).Left = Margin   'reposition the controls and make them visible
        optLineWidth(I).Visible = True
        
        lblLineWidth(I).Left = Margin
        lblLineWidth(I).Height = LineWidths(I)
        lblLineWidth(I).Width = lblColor(0).Width
        lblLineWidth(I).Top = optLineWidth(I).Top + ((optLineWidth(I).Height - LineWidths(I)) / 2)
        lblLineWidth(I).Visible = True
    Next I
End Sub


'''''
''''' Callbacks that implement palette selections
'''''

'Whenever we click on a color label, set the canvas to that color.
'Also, make that label appear selected by giving it a border
Private Sub lblColor_Click(Index As Integer)
    Dim I As Integer
    picCanvas.ForeColor = Colors(Index)
    For I = 0 To MaxColors  'Remove the borders from all the labels
        lblColor(I).BorderStyle = 0
    Next I
    lblColor(Index).BorderStyle = 1 'Add a border to the selected label
End Sub

'When we click on a linewidth label, set the canvas to that width
'Also set the appropriate option button to be the selected one
Private Sub lblLineWidth_Click(Index As Integer)
    picCanvas.DrawWidth = LineWidths(Index)
    optLineWidth(Index) = True
End Sub
'When we click on the LineWidth option, we also set the canvas to the selected line width
Private Sub optLineWidth_Click(Index As Integer)
    picCanvas.DrawWidth = LineWidths(Index)
End Sub

'''''
''''' Callbacks that implement sketching and canvas clearing
'''''

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

'Clear the canvas
Private Sub cmdClear_Click()
    picCanvas.Cls
End Sub
