VERSION 5.00
Object = "{A8B345A0-74B5-11D3-85C2-00105AC8B715}#1.0#0"; "iProfessionalLibrary.ocx"
Object = "{0A362340-2E5E-11D3-85BF-00105AC8B715}#1.0#0"; "isDigitalLibrary.ocx"
Object = "{C5412DA5-2E2F-11D3-85BF-00105AC8B715}#1.0#0"; "isAnalogLibrary.ocx"
Object = "{9F3B4DE1-AA29-11D1-A3D9-FDA4E35D1D25}#1.0#0"; "Io.ocx"
Begin VB.Form PD 
   BackColor       =   &H00200005&
   BorderStyle     =   0  '없음
   Caption         =   "Phase Detector"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "PD.frx":0000
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   Begin iProfessionalLibrary.iSwitchLeverX iSwitchLeverX1 
      Height          =   2295
      Left            =   5040
      TabIndex        =   7
      Top             =   4080
      Width           =   735
      Active          =   -1  'True
      MouseControlStyle=   0
      ShowFocusRect   =   -1  'True
      BackGroundColor =   2097157
      BorderStyle     =   0
      Enabled         =   -1  'True
      Transparent     =   0   'False
      UpdateFrameRate =   60
      OptionSaveAllProperties=   0   'False
      AutoFrameRate   =   0   'False
      Object.Visible         =   -1  'True
      Object.Width           =   49
      Object.Height          =   153
      OPCItemCount    =   0
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '오른쪽 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H00200005&
      BorderStyle     =   0  '없음
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   945
      Left            =   4560
      TabIndex        =   5
      Text            =   "0"
      Top             =   9750
      Width           =   3375
   End
   Begin VB.TextBox txpowertxt 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00200005&
      BorderStyle     =   0  '없음
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   615
      Left            =   8280
      TabIndex        =   4
      Text            =   "0"
      Top             =   8280
      Width           =   1095
   End
   Begin IOLib.IO IO1 
      Left            =   13320
      Top             =   240
      _Version        =   65536
      _ExtentX        =   1270
      _ExtentY        =   1270
      _StockProps     =   0
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2
      Left            =   14160
      Top             =   360
   End
   Begin isDigitalLibrary.iSwitchLedX iSwitchLedX1 
      Height          =   1695
      Left            =   12120
      TabIndex        =   11
      Top             =   2280
      Width           =   2535
      Active          =   0   'False
      ActiveColor     =   65280
      AutoLedSize     =   -1  'True
      Caption         =   "Start"
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionMargin   =   15
      IndicatorAlignment=   1
      IndicatorHeight =   25
      IndicatorMargin =   8
      IndicatorWidth  =   20
      ShowFocusRect   =   -1  'True
      Enabled         =   -1  'True
      BorderStyle     =   0
      Object.Visible         =   -1  'True
      CaptionFontColor=   2097157
      CaptionAlignment=   0
      UpdateFrameRate =   60
      WordWrap        =   0   'False
      Glyph           =   "PD.frx":5355
      BorderSize      =   5
      BorderHighlightColor=   16761024
      BorderShadowColor=   16761024
      BackGroundColor =   8421376
      OptionSaveAllProperties=   0   'False
      AutoFrameRate   =   0   'False
      Object.Width           =   169
      Object.Height          =   113
      MomentaryStyle  =   0
      CaptionFontName =   "Arial"
      CaptionFontSize =   20
      CaptionFontBold =   0   'False
      CaptionFontItalic=   0   'False
      CaptionFontUnderline=   0   'False
      CaptionFontStrikeOut=   0   'False
      OPCItemCount    =   0
   End
   Begin isAnalogLibrary.iLedSpiralX iLedSpiralX1 
      Height          =   1935
      Left            =   12360
      TabIndex        =   12
      Top             =   6360
      Width           =   2175
      SegmentCount    =   32
      SegmentSize     =   3
      SegmentWidth    =   10
      OuterRadius     =   59
      BackGroundColor =   2097157
      BorderStyle     =   0
      SectionColor1   =   65280
      SectionColor2   =   65535
      SectionColor3   =   255
      SectionEnd1     =   50
      SectionEnd2     =   75
      SectionCount    =   1
      ShowOffSegments =   -1  'True
      CurrentMax      =   0
      CurrentMin      =   0
      PositionPercent =   0.26
      Position        =   26
      PositionMax     =   100
      PositionMin     =   0
      Object.Visible         =   -1  'True
      Enabled         =   -1  'True
      BackGroundPicture=   "PD.frx":53AB
      MinMaxFixed     =   0   'False
      Transparent     =   0   'False
      RangeDegrees    =   360
      StartDegrees    =   180
      AutoSize        =   -1  'True
      OuterMargin     =   5
      UpdateFrameRate =   60
      Style           =   1
      OptionSaveAllProperties=   0   'False
      AutoFrameRate   =   0   'False
      Object.Width           =   145
      Object.Height          =   129
      SectionColor4   =   65535
      SectionColor5   =   65535
      SectionEnd3     =   0
      SectionEnd4     =   0
      OPCItemCount    =   0
   End
   Begin iProfessionalLibrary.iSwitchLeverX iSwitchLeverX2 
      Height          =   2295
      Left            =   2280
      TabIndex        =   13
      Top             =   4200
      Width           =   735
      Active          =   0   'False
      MouseControlStyle=   0
      ShowFocusRect   =   -1  'True
      BackGroundColor =   2097157
      BorderStyle     =   0
      Enabled         =   -1  'True
      Transparent     =   0   'False
      UpdateFrameRate =   60
      OptionSaveAllProperties=   0   'False
      AutoFrameRate   =   0   'False
      Object.Visible         =   -1  'True
      Object.Width           =   49
      Object.Height          =   153
      OPCItemCount    =   0
   End
   Begin isAnalogLibrary.iSliderX iSliderX1 
      Height          =   5295
      Left            =   8640
      TabIndex        =   17
      Top             =   2400
      Width           =   1575
      EndsMargin      =   0
      PointerIndicatorInactiveColor=   16711680
      PointerIndicatorActiveColor=   255
      KeyArrowStepSize=   1
      KeyPageStepSize =   10
      Orientation     =   0
      OrientationTickMarks=   0
      PointerHeight   =   15
      PointerStyle    =   3
      PointerWidth    =   30
      ShowFocusRect   =   -1  'True
      TrackColor      =   16744576
      TrackStyle      =   1
      TickMajorStyle  =   0
      TickMinorStyle  =   0
      BackGroundColor =   2097157
      ShowTicksMajor  =   -1  'True
      ShowTicksMinor  =   -1  'True
      ShowTickLabels  =   -1  'True
      TickMajorCount  =   5
      TickMajorColor  =   16777215
      TickMajorLength =   7
      TickMinorAlignment=   1
      TickMinorCount  =   4
      TickMinorColor  =   16777215
      TickMinorLength =   4
      TickMargin      =   6
      TickLabelMargin =   3
      BeginProperty TickLabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TickLabelPrecision=   0
      CurrentMax      =   0
      CurrentMin      =   0
      PositionPercent =   0
      Position        =   0
      PositionMax     =   20
      PositionMin     =   0
      BorderStyle     =   0
      Object.Visible         =   -1  'True
      Enabled         =   -1  'True
      TickLabelFontColor=   16777215
      BackGroundPicture=   "PD.frx":5401
      MinMaxFixed     =   0   'False
      ReverseScale    =   0   'False
      Transparent     =   0   'False
      PrecisionStyle  =   1
      AutoScaleDesiredTicks=   5
      AutoScaleMaxTicks=   5
      AutoScaleEnabled=   0   'False
      AutoScaleStyle  =   1
      MouseControlStyle=   2
      UpdateFrameRate =   60
      OptionSaveAllProperties=   0   'False
      MouseWheelStepSize=   1
      AutoFrameRate   =   0   'False
      Object.Width           =   105
      Object.Height          =   353
      AutoCenter      =   0   'False
      OffsetX         =   0
      OffsetY         =   0
      PointerBitmap   =   "PD.frx":5457
      ShowDisabledState=   -1  'True
      PointerFillEnabled=   -1  'True
      PointerFillColor=   255
      TickLabelFontName=   "Arial"
      TickLabelFontSize=   14
      TickLabelFontBold=   0   'False
      TickLabelFontItalic=   0   'False
      TickLabelFontUnderline=   0   'False
      TickLabelFontStrikeOut=   0   'False
      OPCItemCount    =   0
   End
   Begin VB.Label Label5 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00200005&
      Caption         =   "Tx Level"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   8280
      TabIndex        =   18
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label40 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00200005&
      Caption         =   "Output Port"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   735
      Left            =   2160
      TabIndex        =   16
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label41 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00200005&
      Caption         =   "BNC"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   375
      Left            =   2160
      TabIndex        =   15
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label42 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00200005&
      Caption         =   "PowerLine"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   375
      Left            =   1800
      TabIndex        =   14
      Top             =   7440
      Width           =   1695
   End
   Begin VB.Label Label19 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00200005&
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   12480
      TabIndex        =   10
      Top             =   5160
      Width           =   1935
   End
   Begin VB.Image ImgExit 
      Height          =   1950
      Left            =   13440
      Picture         =   "PD.frx":54AD
      Stretch         =   -1  'True
      Top             =   9600
      Width           =   1920
   End
   Begin VB.Label Label9 
      BackColor       =   &H00200005&
      Caption         =   "Degree"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   8400
      TabIndex        =   9
      Top             =   10080
      Width           =   2055
   End
   Begin VB.Label Label8 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00200005&
      Caption         =   "dBm"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   375
      Left            =   9600
      TabIndex        =   8
      Top             =   8520
      Width           =   615
   End
   Begin VB.Label Label7 
      BackColor       =   &H00200005&
      Caption         =   "Result"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   2640
      TabIndex        =   6
      Top             =   10080
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackColor       =   &H00200005&
      Caption         =   "Receiver"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   375
      Left            =   4800
      TabIndex        =   3
      Top             =   7440
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00200005&
      Caption         =   "Sender"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00200005&
      Caption         =   "Mode"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   4800
      TabIndex        =   1
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00200005&
      Caption         =   "Phase Detector"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   15135
   End
End
Attribute VB_Name = "PD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sigVal As Double
Dim freq As Byte
Dim Amp As Byte
Dim a As Integer
Dim w As Single
Dim t As Single
Dim txCheck As Boolean

Private Sub Form_Load()
Timer1.Enabled = False
sigVal = 0
'If iSwitchRocker3WayX1.Value = 0 Then
'    txCheck = False
'Else: txCheck = True
'End If
End Sub

Private Sub ImgExit_Click()
Unload Me
End Sub

Private Sub iSwitchLedX1_OnChange()
Dim result As Byte
If iSwitchLeverX1.Active = True Then
    If txCheck = False Then
        txCheck = True
        MsgBox "TxPower값이 0입니다.", vbExclamation, "Setting"
        iSwitchLedX1.Active = False
        iSwitchLedX1.Caption = "Start"
        txpowertxt.Text = 0
        iSliderX1.Position = 0
        Timer1.Enabled = False
        iLedSpiralX1.Position = 26
    ElseIf iSliderX1.Position <> 0 Then
        If iSwitchLedX1.Active = True Then
                result = IO1.Open("COM1:", "baud=115200 parity=N data=8 stop=1")
                Timer1.Enabled = True
                t = 0
                sigVal = 0
                iSwitchLedX1.Caption = "Stop"
         ElseIf iSwitchLedX1.Active = False Then
                result = IO1.Close
                Timer1.Enabled = False
                iSwitchLedX1.Caption = "Start"
                iLedSpiralX1.Position = 26
            
        End If
    ElseIf iSliderX1.Position = 0 Then
        txCheck = False
    End If
ElseIf iSwitchLeverX1.Active = False Then
    If iSwitchLedX1.Active = True Then
                result = IO1.Open("COM1:", "baud=115200 parity=N data=8 stop=1")
                Timer1.Enabled = True
                t = 0
                sigVal = 0
                iSwitchLedX1.Caption = "Stop"
         ElseIf iSwitchLedX1.Active = False Then
                result = IO1.Close
                Timer1.Enabled = False
                iSwitchLedX1.Caption = "Start"
                iLedSpiralX1.Position = 26
        End If
End If
End Sub

Private Sub iSwitchLeverX1_OnChange()
    If iSwitchLeverX1.Active = True Then
        Label7.Caption = "Result"
        Label7.ForeColor = vbGreen
        Label9.Caption = "Degree"
        Label9.ForeColor = vbGreen
        Text1.ForeColor = vbGreen
    ElseIf iSwitchLeverX1.Active = False Then
        Label7.Caption = "Result"
        Label7.ForeColor = vbRed
        Label9.Caption = "Degree"
        Label9.ForeColor = vbRed
        Text1.ForeColor = vbRed
    End If
End Sub

Private Sub iSliderX1_OnPositionChangeUser()
a = iSliderX1.Position
If iSwitchLeverX1.Active = True Then
    If a < 0 Then
        txpowertxt.Text = 0
        iSliderX1.Position = 0
        MsgBox "범위를 벗어납니다.", vbExclamation, "Setting"
    ElseIf a = 0 Then
        txCheck = False
        txpowertxt.Text = a
        'Call iSwitchLedX1_OnChange
    ElseIf a > 0 Then
        txpowertxt.Text = a
        txCheck = True
    End If
ElseIf iSwitchLeverX1.Active = False And iSliderX1.Position <> a Then
    MsgBox "Mode를 확인하세요.", vbExclamation, "Setting"
    iSliderX1.Position = a
End If
End Sub

Private Sub Timer1_Timer()
Dim m As Double
Dim pi As Single
    pi = 3.14159
    m = 2 * pi * freq * t 'scal factor 10
    t = t + (Timer1.Interval / 1000)
    sigVal = Amp * 10 * Sin(m)
    IO1.WriteByte (sigVal)
    'Timer1.Enabled = True
    If iLedSpiralX1.Position <> 100 Then
        iLedSpiralX1.Position = iLedSpiralX1.Position + 1
    ElseIf iLedSpiralX1.Position = 100 Then
        iLedSpiralX1.Position = 0
    End If
End Sub
