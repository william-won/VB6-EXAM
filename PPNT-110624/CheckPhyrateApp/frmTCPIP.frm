VERSION 5.00
Begin VB.Form frmTCPIP 
   BorderStyle     =   3  'ũ�� ���� ��ȭ ����
   Caption         =   "TCPIP Settings"
   ClientHeight    =   1665
   ClientLeft      =   2580
   ClientTop       =   1800
   ClientWidth     =   3720
   Icon            =   "frmTCPIP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1665
   ScaleWidth      =   3720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '������ ���
   Begin VB.CheckBox Check1 
      Caption         =   "Trace"
      Height          =   255
      Left            =   2640
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdOKCancel 
      Caption         =   "Cancel"
      Height          =   375
      Index           =   1
      Left            =   2640
      TabIndex        =   3
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton cmdOKCancel 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   2640
      TabIndex        =   2
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox txtPort 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Text            =   "40000"
      Top             =   1080
      Width           =   2295
   End
   Begin VB.TextBox txtRemoteAddress 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "192.168.0.205"
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Remote IP Port"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Remote IP Address"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmTCPIP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()
    If Check1.Value = 1 Then
        frmTelnet.TraceTelnet = True
        frmTelnet.Tracevt100 = True
    Else
        frmTelnet.TraceTelnet = False
        frmTelnet.Tracevt100 = False
    End If
    
End Sub

Private Sub cmdOKCancel_Click(Index As Integer)
On Error Resume Next
    Select Case Index
        Case 0
            
            frmTelnet.WinsockClient.Close
            frmTelnet.WinsockClient.LocalPort = 0
            frmTelnet.RemoteIPAd = txtRemoteAddress
            frmTelnet.RemotePort = txtPort
            frmTelnet.WinsockClient.RemotePort = txtPort
            frmTelnet.WinsockClient.RemoteHost = txtRemoteAddress
            If Err > 0 Then
                MsgBox Error
            Else
                Unload Me
            End If
        Case 1
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
    txtRemoteAddress = frmTelnet.RemoteIPAd
    txtPort = frmTelnet.RemotePort
    Check1.Value = -(frmTelnet.TraceTelnet)
End Sub

Private Sub Label1_Click(Index As Integer)

End Sub
