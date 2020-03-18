VERSION 5.00
Begin VB.Form Navigator 
   BackColor       =   &H00800000&
   BorderStyle     =   0  '없음
   Caption         =   "Form1"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   FillColor       =   &H00800000&
   FillStyle       =   0  '단색
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "Navigator.frx":0000
   ScaleHeight     =   8827.588
   ScaleMode       =   0  '사용자
   ScaleWidth      =   1243.862
   ShowInTaskbar   =   0   'False
   Begin VB.Image imgEM 
      Height          =   1950
      Left            =   5160
      Picture         =   "Navigator.frx":69AF
      Top             =   4800
      Width           =   1800
   End
   Begin VB.Image ImgKS 
      Height          =   1950
      Left            =   2160
      Picture         =   "Navigator.frx":BEE0
      Top             =   7320
      Width           =   1800
   End
   Begin VB.Image imgSnrScope 
      Height          =   1950
      Left            =   11520
      Picture         =   "Navigator.frx":119B9
      Top             =   4800
      Width           =   1800
   End
   Begin VB.Image imgTopology 
      Height          =   1950
      Left            =   8400
      Picture         =   "Navigator.frx":17010
      Top             =   4800
      Width           =   1800
   End
   Begin VB.Image imgGM 
      Height          =   1950
      Left            =   2040
      Picture         =   "Navigator.frx":1C4A3
      Top             =   4815
      Width           =   1800
   End
   Begin VB.Image imgPowerOff 
      Height          =   1950
      Left            =   13560
      Picture         =   "Navigator.frx":22328
      Top             =   9600
      Width           =   1800
   End
   Begin VB.Image imgTerminal 
      Height          =   1950
      Left            =   11400
      Picture         =   "Navigator.frx":26C8B
      Top             =   1680
      Width           =   1800
   End
   Begin VB.Image imgPD 
      Height          =   1950
      Left            =   8280
      Picture         =   "Navigator.frx":2C89B
      Top             =   1680
      Width           =   1800
   End
   Begin VB.Image imgSA 
      Height          =   1950
      Left            =   5160
      Picture         =   "Navigator.frx":32BBD
      Top             =   1680
      Width           =   1800
   End
   Begin VB.Image imgSG 
      Height          =   1950
      Left            =   2040
      Picture         =   "Navigator.frx":387A8
      Top             =   1680
      Width           =   1800
   End
End
Attribute VB_Name = "Navigator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()    'OScilloscope
Dim result As Double
Dim curExec As String
    curExec = App.Path & "\Oscilloscope\Scope.exe"
    result = Shell(curExec, vbNormalFocus)

End Sub

Private Sub Command2_DblClick()    'Prevent double click
    Command2_Click

End Sub

Private Sub imgDLMS_Click()
Dim result As Double
Dim curExec As String
    curExec = App.Path & "\DlmsTester\DlmsTester.exe"
    result = Shell(curExec, vbNormalFocus)

End Sub

Private Sub imgDLMS_DblClick()    'Prevent double click
    imgDLMS_Click

End Sub

Private Sub imgEM_Click()
Dim result As Double
Dim curExec As String
    curExec = App.Path & "\MeterReadSerial\E-TypeMeterReadSerial-v1.0.0.exe"
    result = Shell(curExec, vbMaximizedFocus)

End Sub

Private Sub imgEM_DblClick()    'Prevent double click
    imgEM_Click

End Sub

Private Sub imgGM_Click()
Dim result As Double
Dim curExec As String
    curExec = App.Path & "\MeterReadSerial\G-TypeMeterReadSerial-v1.0.0.exe"
    result = Shell(curExec, vbNormalFocus)

End Sub

Private Sub imgGM_DblClick()    'Prevent double click
    imgGM_Click

End Sub

Private Sub ImgKS_Click()
Dim result As Double
Dim curExec As String
    curExec = App.Path & "\vbpackcap\vbpackcap.exe"
    result = Shell(curExec, vbMaximizedFocus)
End Sub


Private Sub ImgKS_DblClick()    'Prevent double click
    ImgKS_Click

End Sub

Private Sub imgPD_Click()
Dim result As Double
Dim curExec As String
    curExec = App.Path & "\PhaseDetector\PhaseDetector.exe"
    result = Shell(curExec, vbNormalFocus)

End Sub

Private Sub imgPD_DblClick()    'Prevent double click
    imgPD_Click

End Sub


Private Sub imgPowerOff_Click()
    Unload Me
End Sub

Private Sub imgPowerOff_DblClick()    'Prevent double click
    imgPowerOff_Click
    
End Sub

Private Sub imgSA_Click()
Dim result As Double
Dim curExec As String
    curExec = App.Path & "\SpectrumAnalyzer\SpectrumAnalyzer.exe"
    result = Shell(curExec, vbNormalFocus)

End Sub

Private Sub imgSA_DblClick()    'Prevent double click
    imgSA_Click

End Sub

Private Sub imgSnrScope_Click()
Dim result As Double
Dim curExec As String
    curExec = App.Path & "\DS2SnrScope\SNR Scope 2.3.2.exe"
    result = Shell(curExec, vbMaximizedFocus)

End Sub

Private Sub imgSnrScope_DblClick()    'Prevent double click
    imgSnrScope_Click

End Sub

Private Sub imgST_Click()
Dim result As Double
Dim curExec As String
    curExec = App.Path & "\jperf\jperf.bat"
    result = Shell(curExec, vbMaximizedFocus)

End Sub

Private Sub imgST_DblClick()    'Prevent double click
    imgST_Click

End Sub

Private Sub imgTerminal_Click()
Dim result As Double
Dim curExec As String
    curExec = App.Path & "\Terminal\netterm.exe"
    result = Shell(curExec, vbMaximizedFocus)

End Sub

Private Sub imgTerminal_DblClick()    'Prevent double click
    imgTerminal_Click

End Sub

Private Sub imgTopology_Click()
Dim result As Double
Dim curExec As String
    curExec = App.Path & "\DS2Topolgy\Top-1024x768.exe"
    result = Shell(curExec, vbMaximizedFocus)
End Sub

Private Sub imgTopology_DblClick()  'Prevent double click
    imgTopology_Click

End Sub

Private Sub imgSG_Click()
Dim result As Double
Dim curExec As String
    curExec = App.Path & "\SignalGenerator\FG.exe"
    result = Shell(curExec, vbNormalFocus)

End Sub

Private Sub imgSG_DblClick()    'Prevent double click
    imgSG_Click

End Sub

