VERSION 5.00
Begin VB.Form frmSelective 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '단일 고정
   Caption         =   "LP Selective Access"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4920
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.Frame fraValue 
      Caption         =   "Value Selection"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   210
      TabIndex        =   8
      Top             =   1500
      Width           =   4515
      Begin VB.ComboBox cmbFromValue 
         Height          =   330
         Left            =   1320
         Style           =   2  '드롭다운 목록
         TabIndex        =   10
         Top             =   660
         Width           =   975
      End
      Begin VB.ComboBox cmbToValue 
         Height          =   330
         Left            =   3360
         Style           =   2  '드롭다운 목록
         TabIndex        =   9
         Top             =   660
         Width           =   975
      End
      Begin VB.Label lblMaxValue 
         Caption         =   "Max Value :"
         Height          =   285
         Left            =   210
         TabIndex        =   13
         Top             =   330
         Width           =   2235
      End
      Begin VB.Label lblSelect 
         Caption         =   "From Value :"
         Height          =   285
         Index           =   2
         Left            =   210
         TabIndex        =   12
         Top             =   690
         Width           =   1095
      End
      Begin VB.Label lblSelect 
         Caption         =   "To Value :"
         Height          =   285
         Index           =   3
         Left            =   2430
         TabIndex        =   11
         Top             =   690
         Width           =   1095
      End
   End
   Begin VB.Frame fraEntry 
      Caption         =   "Entry Selection"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   210
      TabIndex        =   2
      Top             =   180
      Width           =   4515
      Begin VB.ComboBox cmbToEntry 
         Height          =   330
         Left            =   3360
         Style           =   2  '드롭다운 목록
         TabIndex        =   7
         Top             =   660
         Width           =   975
      End
      Begin VB.ComboBox cmbFromEntry 
         Height          =   330
         Left            =   1320
         Style           =   2  '드롭다운 목록
         TabIndex        =   6
         Top             =   660
         Width           =   975
      End
      Begin VB.Label lblSelect 
         Caption         =   "To Entry :"
         Height          =   285
         Index           =   1
         Left            =   2430
         TabIndex        =   5
         Top             =   690
         Width           =   1095
      End
      Begin VB.Label lblSelect 
         Caption         =   "From Entry :"
         Height          =   285
         Index           =   0
         Left            =   210
         TabIndex        =   4
         Top             =   690
         Width           =   1095
      End
      Begin VB.Label lblMaxEntry 
         Caption         =   "Max Entry :"
         Height          =   285
         Left            =   210
         TabIndex        =   3
         Top             =   330
         Width           =   2235
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3480
      TabIndex        =   1
      Top             =   2880
      Width           =   1245
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   210
      TabIndex        =   0
      Top             =   2880
      Width           =   1245
   End
End
Attribute VB_Name = "frmSelective"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*                                                                                                  *
Dim Now_Doing_Init As Boolean

'*                                                                                                  *

Private Sub Form_Load()
    Call Init_Selective_Form
End Sub
'*                                                                                                  *

Private Sub Init_Selective_Form()
    
    Dim ii_Index As Long
    
    Now_Doing_Init = True
    
    lblMaxEntry.Caption = "Max Entry : " & CStr(Max_LP_Index)
    lblMaxValue.Caption = "Max Value : " & CStr(6)
    
    cmbFromValue.Clear:     cmbToValue.Clear
    For ii_Index = 1 To 6 Step 1
        cmbFromValue.AddItem CStr(ii_Index)
        cmbToValue.AddItem CStr(ii_Index)
    Next ii_Index
    cmbFromValue.ListIndex = 0
    cmbToValue.ListIndex = 5
    
    cmbFromEntry.Clear:     cmbToEntry.Clear
    For ii_Index = 1 To Max_LP_Index Step 1
        cmbFromEntry.AddItem CStr(ii_Index)
        cmbToEntry.AddItem CStr(ii_Index)
    Next ii_Index
    cmbFromEntry.ListIndex = 0
    cmbToEntry.ListIndex = Max_LP_Index - 1
    
    Now_Doing_Init = False
  
End Sub
'*                                                                                                  *

Private Sub cmdOK_Click()
    
    If IsNumeric(cmbFromEntry.Text) And IsNumeric(cmbToEntry.Text) And _
            IsNumeric(cmbFromValue.Text) And IsNumeric(cmbToValue.Text) = False Then
        MsgBox "There is wrong numeric input !", vbExclamation, "Not Numeric Input"
        Exit Sub
    End If
    If CLng(cmbFromEntry.Text) > CLng(cmbToEntry.Text) Then
        MsgBox "Wrong Entry Item Input ! (From > To)", vbExclamation, "Wrong Entry"
        Exit Sub
    End If
    If CLng(cmbFromValue.Text) > CLng(cmbToValue.Text) Then
        MsgBox "Wrong Value Item Input ! (From > To)", vbExclamation, "Wrong Value"
        Exit Sub
    End If
    
    With sAccess
        .fromEntry = CLng(cmbFromEntry.Text)
        .toEntry = CLng(cmbToEntry.Text)
        .fromValue = CLng(cmbFromValue.Text)
        .toValue = CLng(cmbToValue.Text)
    End With
    Selective_Or_Not = True
    
    Unload Me
    
End Sub
'*                                                                                                  *

Private Sub cmdCancel_Click()

    With sAccess
        .fromEntry = 0
        .toEntry = 0
        .fromValue = 0
        .toValue = 0
    End With
    Selective_Or_Not = False

    Unload Me
    
End Sub
'*                                                                                                  *

Private Sub cmbFromEntry_Click()

    Dim ii_Index As Long
    
    If Now_Doing_Init = True Then Exit Sub
    
    cmbToEntry.Clear
    For ii_Index = cmbFromEntry.Text To Max_LP_Index Step 1
        cmbToEntry.AddItem CStr(ii_Index)
    Next ii_Index
    cmbToEntry.ListIndex = 0
    
End Sub
'*                                                                                                  *

Private Sub cmbFromValue_Click()

    Dim ii_Index As Long

    If Now_Doing_Init = True Then Exit Sub

    cmbToValue.Clear
    For ii_Index = cmbFromValue.Text To 6 Step 1
        cmbToValue.AddItem CStr(ii_Index)
    Next ii_Index
    cmbToValue.ListIndex = cmbToValue.ListCount - 1

End Sub
'*                                                                                                  *


