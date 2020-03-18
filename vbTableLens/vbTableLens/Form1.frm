VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Table Lens Example"
   ClientHeight    =   7920
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6000
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7920
   ScaleWidth      =   6000
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExpandAll 
      Caption         =   "Expand All"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3660
      TabIndex        =   2
      Top             =   720
      Width           =   1995
   End
   Begin VB.CommandButton cmdCollapseAll 
      Caption         =   "Collapse All"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3660
      TabIndex        =   1
      Top             =   120
      Width           =   1995
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   3660
      TabIndex        =   3
      Top             =   1440
      Width           =   1995
   End
   Begin VB.Label lblCell 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'A very simple demonstration of a 'table lens' where a user can
'view a single column of values as either a graphical bar or its numeric value.

Option Explicit
Const MaxCells As Integer = 100 'The number of cells in the table column
Dim Table(MaxCells) As Integer

'Initialize the table
Private Sub Form_Load()
    Dim I As Integer
    Form1.ScaleMode = vbPixels    'Measures are all in pixels
    Label1.Caption = "Click on a cell to expand/collapse it, " & _
                     "or on the buttons to expand/collapse all cells. " & _
                     "Hover over a cell to see its index and value."
                     
    CreateTable                   'Now create the Table
End Sub

'Create the Table of values and associated labels
' (which will become the 'cells' of the table)
Private Sub CreateTable()
    Dim I As Integer
    Const Max As Integer = 100
    Const Min As Integer = 1
    
    'Create a table of random values between Max and Min
    For I = 0 To MaxCells
        Table(I) = Int((Max - Min + 1) * Rnd) + Min
    Next I
    
    'Now create the cells.
    For I = 0 To MaxCells           'then we create new labels corresponding to the other colors.
        If I <> 0 Then Load lblCell(I)
        lblCell(I) = ""
        lblCell(I).ToolTipText = "Cell " & I & ", Value = " & Table(I)
    Next I
    TableDraw
End Sub

'Draw the table of cells.
'If the caption field is empty, we draw it as a bar
'If it contains a value, we draw it as a numeric value
Private Sub TableDraw()
    Const Margin As Integer = 10    'The left/top margin
    Const BarHeight As Integer = 4  'The height of the graphical bar
    Const CellPadding = 1            'The vertical padding (spacing) between cells
    Dim I As Integer
    
    For I = 0 To MaxCells
        'If there is no caption contents, then draw it as a bar
        If lblCell(I).Caption = "" Then
            lblCell(I).AutoSize = False
            lblCell(I).Height = BarHeight
            lblCell(I).Width = Table(I)
            lblCell(I).BackColor = vbRed
        'Otherwise draw it as a cell
        Else
            lblCell(I).AutoSize = True
            lblCell(I).BackColor = vbWhite
        End If
        
        'Now position the cell on the form
        If I = 0 Then
            lblCell(I).Top = Margin
        Else
            lblCell(I).Top = lblCell(I - 1).Top + lblCell(I - 1).Height + CellPadding
        End If
        lblCell(I).Left = Margin
        lblCell(I).Visible = True
    Next I
    Form1.ScaleMode = vbTwips
    Form1.Height = lblCell(MaxCells).Top + lblCell(MaxCells).Height + (2 * Margin) + 500
    Form1.ScaleMode = vbPixels
End Sub


'''
''' CALLBACKS
'''
'Collapse all cells into its graphical representation
Private Sub cmdCollapseAll_Click()
    Dim I As Integer
    For I = 0 To MaxCells
        lblCell(I).Caption = ""
    Next I
    TableDraw
End Sub

'Expand all cells into its numeric representation
Private Sub cmdExpandAll_Click()
    Dim I As Integer
    For I = 0 To MaxCells
        lblCell(I).Caption = Table(I)
    Next I
    TableDraw
End Sub


'A user clicked on a cell: toggle its collapsed/expanded state.
Private Sub lblCell_Click(Index As Integer)
    If lblCell(Index).Caption = "" Then
        lblCell(Index).Caption = Table(Index)
    Else
        lblCell(Index).Caption = ""
    End If
    TableDraw
End Sub

