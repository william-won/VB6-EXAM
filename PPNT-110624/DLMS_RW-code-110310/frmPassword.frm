VERSION 5.00
Begin VB.Form frmPassword 
   BorderStyle     =   1  '단일 고정
   Caption         =   "Password Setting"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4215
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   4215
   StartUpPosition =   1  '소유자 가운데
   Begin VB.TextBox txtPassword 
      Alignment       =   2  '가운데 맞춤
      Height          =   300
      Index           =   2
      Left            =   2700
      MaxLength       =   8
      TabIndex        =   7
      Text            =   "1A2B3C4D"
      Top             =   1050
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.TextBox txtPassword 
      Alignment       =   2  '가운데 맞춤
      Height          =   300
      Index           =   1
      Left            =   2700
      MaxLength       =   8
      TabIndex        =   6
      Text            =   "1A2B3C4D"
      Top             =   630
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.TextBox txtPassword 
      Alignment       =   2  '가운데 맞춤
      Height          =   300
      Index           =   0
      Left            =   2700
      MaxLength       =   8
      TabIndex        =   5
      Text            =   "1A2B3C4D"
      Top             =   210
      Width           =   1305
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   525
      Left            =   2400
      TabIndex        =   1
      Top             =   1560
      Width           =   1635
   End
   Begin VB.CommandButton cmdSet 
      Caption         =   "&Save Password"
      Default         =   -1  'True
      Height          =   525
      Left            =   180
      TabIndex        =   0
      Top             =   1560
      Width           =   1635
   End
   Begin VB.Label lblTitle 
      Caption         =   "Master Password :"
      Height          =   285
      Index           =   2
      Left            =   150
      TabIndex        =   4
      Top             =   1050
      Visible         =   0   'False
      Width           =   2385
   End
   Begin VB.Label lblTitle 
      Caption         =   "Factory Password :"
      Height          =   285
      Index           =   1
      Left            =   150
      TabIndex        =   3
      Top             =   630
      Visible         =   0   'False
      Width           =   2385
   End
   Begin VB.Label lblTitle 
      Caption         =   "Manufacturer Password :"
      Height          =   285
      Index           =   0
      Left            =   150
      TabIndex        =   2
      Top             =   210
      Width           =   2385
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*                                                                                                  *

Private Sub Form_Load()

    txtPassword(0).Text = "1A2B3C4D"
    txtPassword(1).Text = "1A2B3C4D"
    txtPassword(2).Text = "1A2B3C4D"

End Sub
'*                                                                                                  *

Private Sub cmdCancel_Click()
    
    Unload Me
    
End Sub
'*                                                                                                  *

Private Sub cmdSet_Click()
    
    Dim ii_Byte As Byte
    
    For ii_Byte = 0 To 2 Step 1
        If Len(txtPassword(ii_Byte).Text) > 8 Then
            txtPassword(ii_Byte).Text = Mid(txtPassword(ii_Byte).Text, 1, 8)
        ElseIf Len(txtPassword(ii_Byte).Text) < 8 Then
            txtPassword(ii_Byte).Text = txtPassword(ii_Byte).Text & _
                                        String(8 - Len(txtPassword(ii_Byte).Text), " ")
        End If
    Next ii_Byte
    
    gPASS_Manufacture = txtPassword(0).Text
    gPASS_Factory = txtPassword(1).Text
    gPASS_Master = txtPassword(2).Text
    
    Unload Me
    
End Sub
'*                                                                                                  *

