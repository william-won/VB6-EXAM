VERSION 5.00
Begin VB.Form MoveListForm 
   Caption         =   "MoveList"
   ClientHeight    =   3465
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3465
   ScaleWidth      =   4800
   Begin VB.CommandButton cmdMoveLeft 
      Height          =   375
      Left            =   2160
      Picture         =   "MoveList.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton cmdMoveRight 
      Height          =   375
      Left            =   2160
      Picture         =   "MoveList.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   360
      Width           =   495
   End
   Begin VB.ListBox lstRight 
      Height          =   2400
      Left            =   2760
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   180
      Width           =   1935
   End
   Begin VB.ListBox lstLeft 
      Height          =   2400
      Left            =   120
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label lblHelp 
      Caption         =   "Label1"
      Height          =   675
      Left            =   120
      TabIndex        =   4
      Top             =   2700
      Width           =   4575
   End
End
Attribute VB_Name = "MoveListForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This (slightly modified) example came from
'  Rod Stephens "Ready-to-run Visual Basic Code Library", Wiley Press
'If you can get a copy of this book, it has lots of great examples and tips.

'This program illustrates how to use listboxes,
'Specifically, its shows have 1 or more items can be moved between two lists.
'It also how to add images to buttons (see the button Picture and Style Properties)
Option Explicit

'Create and add items to the list
Private Sub Form_Load()
    ' Put some values in the left list.
    lstLeft.AddItem "Ape"
    lstLeft.AddItem "Bear"
    lstLeft.AddItem "Cat"
    lstLeft.AddItem "Dog"
    lstLeft.AddItem "Eagle"
    lstLeft.AddItem "Frog"
    lstLeft.AddItem "Giraffe"
    lstLeft.AddItem "Hen"
    lstLeft.AddItem "Ibex"
    lstLeft.AddItem "Jackel"
    lblHelp.WordWrap = True
    lblHelp.Caption = "Select one item by Clicking it." & vbCrLf & _
                      "Select several contiguous items by Shift-clicking." & vbCrLf & _
                      "Select several non-contiguous items by Control-clicking."
                      
    ' Enable the appropriate buttons.
    EnableButtons
End Sub

' Enable / Disabel the appropriate buttons.
Private Sub EnableButtons()
Dim i As Integer

    ' See if an item is selected in the left list.
    For i = lstLeft.ListCount - 1 To 0 Step -1
        If lstLeft.Selected(i) Then Exit For
    Next i
    cmdMoveRight.Enabled = (i >= 0)

    ' See if an item is selected in the right list.
    For i = lstRight.ListCount - 1 To 0 Step -1
        If lstRight.Selected(i) Then Exit For
    Next i
    cmdMoveLeft.Enabled = (i >= 0)
End Sub

' Move the items selected in the right list into the left list.
Private Sub cmdMoveLeft_Click()
Dim i As Integer

    ' Remove the selected items.
    For i = lstRight.ListCount - 1 To 0 Step -1
        ' Move this item.
        If lstRight.Selected(i) Then
            lstLeft.AddItem lstRight.List(i)
            lstRight.RemoveItem i
        End If
    Next i

    ' Enable the correct buttons.
    EnableButtons
End Sub

' Move the items selected in the left list into the right list.
Private Sub cmdMoveRight_Click()
Dim i As Integer

    ' Remove the selected items.
    For i = lstLeft.ListCount - 1 To 0 Step -1
        ' Move this item.
        If lstLeft.Selected(i) Then
            lstRight.AddItem lstLeft.List(i)
            lstLeft.RemoveItem i
        End If
    Next i

    ' Enable the correct buttons.
    EnableButtons
End Sub


' Enable the appropriate buttons.
Private Sub lstLeft_Click()
    EnableButtons
End Sub
' Enable the appropriate buttons.
Private Sub lstRight_Click()
    EnableButtons
End Sub

