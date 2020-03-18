VERSION 5.00
Object = "{A8B345A0-74B5-11D3-85C2-00105AC8B715}#1.0#0"; "iProfessionalLibrary.ocx"
Object = "{0A362340-2E5E-11D3-85BF-00105AC8B715}#1.0#0"; "isDigitalLibrary.ocx"
Object = "{C5412DA5-2E2F-11D3-85BF-00105AC8B715}#1.0#0"; "isAnalogLibrary.ocx"
Object = "{9F3B4DE1-AA29-11D1-A3D9-FDA4E35D1D25}#1.0#0"; "Io.ocx"
Begin VB.Form SigGen 
   BackColor       =   &H00200005&
   BorderStyle     =   0  '없음
   Caption         =   "&H00FFFFFF&"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "SigGen.frx":0000
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   Begin isDigitalLibrary.iSwitchLedX iSwitchLedX2 
      Height          =   1335
      Left            =   2880
      TabIndex        =   61
      Top             =   2400
      Width           =   1095
      Active          =   -1  'True
      ActiveColor     =   255
      AutoLedSize     =   -1  'True
      Caption         =   "On"
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionMargin   =   10
      IndicatorAlignment=   1
      IndicatorHeight =   11
      IndicatorMargin =   5
      IndicatorWidth  =   10
      ShowFocusRect   =   -1  'True
      Enabled         =   -1  'True
      BorderStyle     =   0
      Object.Visible         =   -1  'True
      CaptionFontColor=   0
      CaptionAlignment=   0
      UpdateFrameRate =   60
      WordWrap        =   0   'False
      Glyph           =   "SigGen.frx":5355
      BorderSize      =   2
      BorderHighlightColor=   -16777196
      BorderShadowColor=   2097157
      BackGroundColor =   12632256
      OptionSaveAllProperties=   0   'False
      AutoFrameRate   =   0   'False
      Object.Width           =   73
      Object.Height          =   89
      MomentaryStyle  =   0
      CaptionFontName =   "Arial"
      CaptionFontSize =   14
      CaptionFontBold =   0   'False
      CaptionFontItalic=   0   'False
      CaptionFontUnderline=   0   'False
      CaptionFontStrikeOut=   0   'False
      OPCItemCount    =   0
   End
   Begin iProfessionalLibrary.iSwitchLeverX iSwitchLeverX1 
      Height          =   2295
      Left            =   720
      TabIndex        =   57
      Top             =   4440
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
   Begin isDigitalLibrary.iSwitchLedX iSwitchLedX1 
      Height          =   1695
      Left            =   12240
      TabIndex        =   56
      Top             =   1680
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
      Glyph           =   "SigGen.frx":53AB
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
   Begin VB.TextBox dBtxt 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00200005&
      BorderStyle     =   0  '없음
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   735
      Left            =   9840
      TabIndex        =   26
      Text            =   "0"
      Top             =   8040
      Width           =   975
   End
   Begin VB.TextBox Freq5txt 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00200005&
      BorderStyle     =   0  '없음
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   810
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "0"
      Top             =   8880
      Width           =   735
   End
   Begin VB.TextBox Freq4txt 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00200005&
      BorderStyle     =   0  '없음
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   810
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   "0"
      Top             =   7320
      Width           =   735
   End
   Begin VB.TextBox Freq3txt 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00200005&
      BorderStyle     =   0  '없음
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   810
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "0"
      Top             =   5760
      Width           =   735
   End
   Begin VB.TextBox Freq2txt 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00200005&
      BorderStyle     =   0  '없음
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   810
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "0"
      Top             =   4200
      Width           =   735
   End
   Begin iProfessionalLibrary.iSwitchRocker3WayX iSwitchRocker3WayX1 
      Height          =   1335
      Left            =   7800
      TabIndex        =   6
      Top             =   2400
      Width           =   1095
      ShowFocusRect   =   -1  'True
      BorderMargin    =   2
      RepeatInitialDelay=   500
      RepeatInterval  =   50
      Value           =   0
      Increment       =   10
      UseArrowKeys    =   -1  'True
      BackGroundColor =   16761024
      BorderStyle     =   0
      Enabled         =   -1  'True
      Transparent     =   0   'False
      UpdateFrameRate =   60
      OptionSaveAllProperties=   0   'False
      AutoFrameRate   =   0   'False
      Object.Visible         =   -1  'True
      Object.Width           =   73
      Object.Height          =   89
      OPCItemCount    =   0
   End
   Begin VB.TextBox Freq1txt 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00200005&
      BorderStyle     =   0  '없음
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   810
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "1"
      Top             =   2640
      Width           =   735
   End
   Begin isAnalogLibrary.iSliderX iSliderX1 
      Height          =   5295
      Left            =   9840
      TabIndex        =   3
      Top             =   2520
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
      BackGroundPicture=   "SigGen.frx":5401
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
      PointerBitmap   =   "SigGen.frx":5457
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
   Begin IOLib.IO IO1 
      Left            =   14640
      Top             =   600
      _Version        =   65536
      _ExtentX        =   1270
      _ExtentY        =   1270
      _StockProps     =   0
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2
      Left            =   14880
      Top             =   120
   End
   Begin iProfessionalLibrary.iSwitchRocker3WayX iSwitchRocker3WayX2 
      Height          =   1335
      Left            =   7800
      TabIndex        =   9
      Top             =   3960
      Width           =   1095
      ShowFocusRect   =   -1  'True
      BorderMargin    =   2
      RepeatInitialDelay=   500
      RepeatInterval  =   50
      Value           =   0
      Increment       =   10
      UseArrowKeys    =   -1  'True
      BackGroundColor =   16761024
      BorderStyle     =   0
      Enabled         =   -1  'True
      Transparent     =   0   'False
      UpdateFrameRate =   60
      OptionSaveAllProperties=   0   'False
      AutoFrameRate   =   0   'False
      Object.Visible         =   -1  'True
      Object.Width           =   73
      Object.Height          =   89
      OPCItemCount    =   0
   End
   Begin iProfessionalLibrary.iSwitchRocker3WayX iSwitchRocker3WayX3 
      Height          =   1335
      Left            =   7800
      TabIndex        =   12
      Top             =   5520
      Width           =   1095
      ShowFocusRect   =   -1  'True
      BorderMargin    =   2
      RepeatInitialDelay=   500
      RepeatInterval  =   50
      Value           =   0
      Increment       =   10
      UseArrowKeys    =   -1  'True
      BackGroundColor =   16761024
      BorderStyle     =   0
      Enabled         =   -1  'True
      Transparent     =   0   'False
      UpdateFrameRate =   60
      OptionSaveAllProperties=   0   'False
      AutoFrameRate   =   0   'False
      Object.Visible         =   -1  'True
      Object.Width           =   73
      Object.Height          =   89
      OPCItemCount    =   0
   End
   Begin iProfessionalLibrary.iSwitchRocker3WayX iSwitchRocker3WayX4 
      Height          =   1335
      Left            =   7800
      TabIndex        =   15
      Top             =   7080
      Width           =   1095
      ShowFocusRect   =   -1  'True
      BorderMargin    =   2
      RepeatInitialDelay=   500
      RepeatInterval  =   50
      Value           =   0
      Increment       =   10
      UseArrowKeys    =   -1  'True
      BackGroundColor =   16761024
      BorderStyle     =   0
      Enabled         =   -1  'True
      Transparent     =   0   'False
      UpdateFrameRate =   60
      OptionSaveAllProperties=   0   'False
      AutoFrameRate   =   0   'False
      Object.Visible         =   -1  'True
      Object.Width           =   73
      Object.Height          =   89
      OPCItemCount    =   0
   End
   Begin iProfessionalLibrary.iSwitchRocker3WayX iSwitchRocker3WayX5 
      Height          =   1335
      Left            =   7800
      TabIndex        =   18
      Top             =   8640
      Width           =   1095
      ShowFocusRect   =   -1  'True
      BorderMargin    =   2
      RepeatInitialDelay=   500
      RepeatInterval  =   50
      Value           =   0
      Increment       =   10
      UseArrowKeys    =   -1  'True
      BackGroundColor =   16761024
      BorderStyle     =   0
      Enabled         =   -1  'True
      Transparent     =   0   'False
      UpdateFrameRate =   60
      OptionSaveAllProperties=   0   'False
      AutoFrameRate   =   0   'False
      Object.Visible         =   -1  'True
      Object.Width           =   73
      Object.Height          =   89
      OPCItemCount    =   0
   End
   Begin iProfessionalLibrary.iSwitchRocker3WayX iSwitchRocker3WayX7 
      Height          =   1335
      Left            =   6000
      TabIndex        =   19
      Top             =   8640
      Width           =   1095
      ShowFocusRect   =   -1  'True
      BorderMargin    =   2
      RepeatInitialDelay=   500
      RepeatInterval  =   50
      Value           =   0
      Increment       =   1
      UseArrowKeys    =   -1  'True
      BackGroundColor =   16761024
      BorderStyle     =   0
      Enabled         =   -1  'True
      Transparent     =   0   'False
      UpdateFrameRate =   60
      OptionSaveAllProperties=   0   'False
      AutoFrameRate   =   0   'False
      Object.Visible         =   -1  'True
      Object.Width           =   73
      Object.Height          =   89
      OPCItemCount    =   0
   End
   Begin iProfessionalLibrary.iSwitchRocker3WayX iSwitchRocker3WayX8 
      Height          =   1335
      Left            =   6000
      TabIndex        =   20
      Top             =   7080
      Width           =   1095
      ShowFocusRect   =   -1  'True
      BorderMargin    =   2
      RepeatInitialDelay=   500
      RepeatInterval  =   50
      Value           =   0
      Increment       =   1
      UseArrowKeys    =   -1  'True
      BackGroundColor =   16761024
      BorderStyle     =   0
      Enabled         =   -1  'True
      Transparent     =   0   'False
      UpdateFrameRate =   60
      OptionSaveAllProperties=   0   'False
      AutoFrameRate   =   0   'False
      Object.Visible         =   -1  'True
      Object.Width           =   73
      Object.Height          =   89
      OPCItemCount    =   0
   End
   Begin iProfessionalLibrary.iSwitchRocker3WayX iSwitchRocker3WayX9 
      Height          =   1335
      Left            =   6000
      TabIndex        =   21
      Top             =   5520
      Width           =   1095
      ShowFocusRect   =   -1  'True
      BorderMargin    =   2
      RepeatInitialDelay=   500
      RepeatInterval  =   50
      Value           =   0
      Increment       =   1
      UseArrowKeys    =   -1  'True
      BackGroundColor =   16761024
      BorderStyle     =   0
      Enabled         =   -1  'True
      Transparent     =   0   'False
      UpdateFrameRate =   60
      OptionSaveAllProperties=   0   'False
      AutoFrameRate   =   0   'False
      Object.Visible         =   -1  'True
      Object.Width           =   73
      Object.Height          =   89
      OPCItemCount    =   0
   End
   Begin iProfessionalLibrary.iSwitchRocker3WayX iSwitchRocker3WayX10 
      Height          =   1335
      Left            =   6000
      TabIndex        =   22
      Top             =   3960
      Width           =   1095
      ShowFocusRect   =   -1  'True
      BorderMargin    =   2
      RepeatInitialDelay=   500
      RepeatInterval  =   50
      Value           =   0
      Increment       =   1
      UseArrowKeys    =   -1  'True
      BackGroundColor =   16761024
      BorderStyle     =   0
      Enabled         =   -1  'True
      Transparent     =   0   'False
      UpdateFrameRate =   60
      OptionSaveAllProperties=   0   'False
      AutoFrameRate   =   0   'False
      Object.Visible         =   -1  'True
      Object.Width           =   73
      Object.Height          =   89
      OPCItemCount    =   0
   End
   Begin iProfessionalLibrary.iSwitchRocker3WayX iSwitchRocker3WayX11 
      Height          =   1335
      Left            =   6000
      TabIndex        =   23
      Top             =   2400
      Width           =   1095
      ShowFocusRect   =   -1  'True
      BorderMargin    =   2
      RepeatInitialDelay=   500
      RepeatInterval  =   50
      Value           =   0
      Increment       =   1
      UseArrowKeys    =   -1  'True
      BackGroundColor =   16761024
      BorderStyle     =   0
      Enabled         =   -1  'True
      Transparent     =   0   'False
      UpdateFrameRate =   60
      OptionSaveAllProperties=   0   'False
      AutoFrameRate   =   0   'False
      Object.Visible         =   -1  'True
      Object.Width           =   73
      Object.Height          =   89
      OPCItemCount    =   0
   End
   Begin isAnalogLibrary.iLedSpiralX iLedSpiralX1 
      Height          =   1935
      Left            =   12360
      TabIndex        =   55
      Top             =   5280
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
      CurrentMax      =   26
      CurrentMin      =   26
      PositionPercent =   0.26
      Position        =   26
      PositionMax     =   100
      PositionMin     =   0
      Object.Visible         =   -1  'True
      Enabled         =   -1  'True
      BackGroundPicture=   "SigGen.frx":54AD
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
   Begin isDigitalLibrary.iSwitchLedX iSwitchLedX3 
      Height          =   1335
      Left            =   2880
      TabIndex        =   62
      Top             =   3960
      Width           =   1095
      Active          =   0   'False
      ActiveColor     =   255
      AutoLedSize     =   -1  'True
      Caption         =   "On"
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionMargin   =   10
      IndicatorAlignment=   1
      IndicatorHeight =   11
      IndicatorMargin =   5
      IndicatorWidth  =   10
      ShowFocusRect   =   -1  'True
      Enabled         =   -1  'True
      BorderStyle     =   0
      Object.Visible         =   -1  'True
      CaptionFontColor=   0
      CaptionAlignment=   0
      UpdateFrameRate =   60
      WordWrap        =   0   'False
      Glyph           =   "SigGen.frx":5503
      BorderSize      =   2
      BorderHighlightColor=   -16777196
      BorderShadowColor=   2097157
      BackGroundColor =   12632256
      OptionSaveAllProperties=   0   'False
      AutoFrameRate   =   0   'False
      Object.Width           =   73
      Object.Height          =   89
      MomentaryStyle  =   0
      CaptionFontName =   "Arial"
      CaptionFontSize =   14
      CaptionFontBold =   0   'False
      CaptionFontItalic=   0   'False
      CaptionFontUnderline=   0   'False
      CaptionFontStrikeOut=   0   'False
      OPCItemCount    =   0
   End
   Begin isDigitalLibrary.iSwitchLedX iSwitchLedX4 
      Height          =   1335
      Left            =   2880
      TabIndex        =   63
      Top             =   5520
      Width           =   1095
      Active          =   0   'False
      ActiveColor     =   255
      AutoLedSize     =   -1  'True
      Caption         =   "On"
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionMargin   =   10
      IndicatorAlignment=   1
      IndicatorHeight =   11
      IndicatorMargin =   5
      IndicatorWidth  =   10
      ShowFocusRect   =   -1  'True
      Enabled         =   -1  'True
      BorderStyle     =   0
      Object.Visible         =   -1  'True
      CaptionFontColor=   0
      CaptionAlignment=   0
      UpdateFrameRate =   60
      WordWrap        =   0   'False
      Glyph           =   "SigGen.frx":5559
      BorderSize      =   2
      BorderHighlightColor=   -16777196
      BorderShadowColor=   2097157
      BackGroundColor =   12632256
      OptionSaveAllProperties=   0   'False
      AutoFrameRate   =   0   'False
      Object.Width           =   73
      Object.Height          =   89
      MomentaryStyle  =   0
      CaptionFontName =   "Arial"
      CaptionFontSize =   14
      CaptionFontBold =   0   'False
      CaptionFontItalic=   0   'False
      CaptionFontUnderline=   0   'False
      CaptionFontStrikeOut=   0   'False
      OPCItemCount    =   0
   End
   Begin isDigitalLibrary.iSwitchLedX iSwitchLedX5 
      Height          =   1335
      Left            =   2880
      TabIndex        =   64
      Top             =   7080
      Width           =   1095
      Active          =   0   'False
      ActiveColor     =   255
      AutoLedSize     =   -1  'True
      Caption         =   "On"
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionMargin   =   10
      IndicatorAlignment=   1
      IndicatorHeight =   11
      IndicatorMargin =   5
      IndicatorWidth  =   10
      ShowFocusRect   =   -1  'True
      Enabled         =   -1  'True
      BorderStyle     =   0
      Object.Visible         =   -1  'True
      CaptionFontColor=   0
      CaptionAlignment=   0
      UpdateFrameRate =   60
      WordWrap        =   0   'False
      Glyph           =   "SigGen.frx":55AF
      BorderSize      =   2
      BorderHighlightColor=   -16777196
      BorderShadowColor=   2097157
      BackGroundColor =   12632256
      OptionSaveAllProperties=   0   'False
      AutoFrameRate   =   0   'False
      Object.Width           =   73
      Object.Height          =   89
      MomentaryStyle  =   0
      CaptionFontName =   "Arial"
      CaptionFontSize =   14
      CaptionFontBold =   0   'False
      CaptionFontItalic=   0   'False
      CaptionFontUnderline=   0   'False
      CaptionFontStrikeOut=   0   'False
      OPCItemCount    =   0
   End
   Begin isDigitalLibrary.iSwitchLedX iSwitchLedX6 
      Height          =   1335
      Left            =   2880
      TabIndex        =   65
      Top             =   8640
      Width           =   1095
      Active          =   0   'False
      ActiveColor     =   255
      AutoLedSize     =   -1  'True
      Caption         =   "On"
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionMargin   =   10
      IndicatorAlignment=   1
      IndicatorHeight =   11
      IndicatorMargin =   5
      IndicatorWidth  =   10
      ShowFocusRect   =   -1  'True
      Enabled         =   -1  'True
      BorderStyle     =   0
      Object.Visible         =   -1  'True
      CaptionFontColor=   0
      CaptionAlignment=   0
      UpdateFrameRate =   60
      WordWrap        =   0   'False
      Glyph           =   "SigGen.frx":5605
      BorderSize      =   2
      BorderHighlightColor=   -16777196
      BorderShadowColor=   2097157
      BackGroundColor =   12632256
      OptionSaveAllProperties=   0   'False
      AutoFrameRate   =   0   'False
      Object.Width           =   73
      Object.Height          =   89
      MomentaryStyle  =   0
      CaptionFontName =   "Arial"
      CaptionFontSize =   14
      CaptionFontBold =   0   'False
      CaptionFontItalic=   0   'False
      CaptionFontUnderline=   0   'False
      CaptionFontStrikeOut=   0   'False
      OPCItemCount    =   0
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
      Height          =   375
      Left            =   12480
      TabIndex        =   34
      Top             =   4680
      Width           =   1935
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
      Left            =   240
      TabIndex        =   60
      Top             =   6960
      Width           =   1695
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
      Left            =   600
      TabIndex        =   59
      Top             =   3840
      Width           =   975
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
      Left            =   600
      TabIndex        =   58
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label39 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00200005&
      Caption         =   "Up"
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
      Left            =   7080
      TabIndex        =   54
      Top             =   8640
      Width           =   615
   End
   Begin VB.Label Label38 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00200005&
      Caption         =   "Dn"
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
      Left            =   7080
      TabIndex        =   53
      Top             =   9600
      Width           =   615
   End
   Begin VB.Label Label37 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00200005&
      Caption         =   "Up"
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
      Left            =   7080
      TabIndex        =   52
      Top             =   7080
      Width           =   615
   End
   Begin VB.Label Label36 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00200005&
      Caption         =   "Dn"
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
      Left            =   7080
      TabIndex        =   51
      Top             =   8040
      Width           =   615
   End
   Begin VB.Label Label35 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00200005&
      Caption         =   "Up"
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
      Left            =   7080
      TabIndex        =   50
      Top             =   5520
      Width           =   615
   End
   Begin VB.Label Label34 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00200005&
      Caption         =   "Dn"
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
      Left            =   7080
      TabIndex        =   49
      Top             =   6480
      Width           =   615
   End
   Begin VB.Label Label33 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00200005&
      Caption         =   "Up"
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
      Left            =   7080
      TabIndex        =   48
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label Label32 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00200005&
      Caption         =   "Dn"
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
      Left            =   7080
      TabIndex        =   47
      Top             =   4920
      Width           =   615
   End
   Begin VB.Label Label31 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00200005&
      Caption         =   "Up"
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
      Left            =   7080
      TabIndex        =   46
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label Label30 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00200005&
      Caption         =   "Dn"
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
      Left            =   7080
      TabIndex        =   45
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label Label29 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00200005&
      Caption         =   "Up"
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
      Left            =   8880
      TabIndex        =   44
      Top             =   8640
      Width           =   615
   End
   Begin VB.Label Label28 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00200005&
      Caption         =   "Dn"
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
      Left            =   8880
      TabIndex        =   43
      Top             =   9600
      Width           =   615
   End
   Begin VB.Label Label27 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00200005&
      Caption         =   "Up"
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
      Left            =   8880
      TabIndex        =   42
      Top             =   7080
      Width           =   615
   End
   Begin VB.Label Label26 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00200005&
      Caption         =   "Dn"
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
      Left            =   8880
      TabIndex        =   41
      Top             =   8040
      Width           =   615
   End
   Begin VB.Label Label25 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00200005&
      Caption         =   "Up"
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
      Left            =   8880
      TabIndex        =   40
      Top             =   5520
      Width           =   615
   End
   Begin VB.Label Label24 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00200005&
      Caption         =   "Dn"
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
      Left            =   8880
      TabIndex        =   39
      Top             =   6480
      Width           =   615
   End
   Begin VB.Label Label23 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00200005&
      Caption         =   "Up"
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
      Left            =   8880
      TabIndex        =   38
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label Label22 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00200005&
      Caption         =   "Dn"
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
      Left            =   8880
      TabIndex        =   37
      Top             =   4920
      Width           =   615
   End
   Begin VB.Label Label21 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00200005&
      Caption         =   "Up"
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
      Left            =   8880
      TabIndex        =   36
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label Label20 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00200005&
      Caption         =   "Dn"
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
      Left            =   8880
      TabIndex        =   35
      Top             =   3360
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   1950
      Left            =   13440
      Picture         =   "SigGen.frx":565B
      Stretch         =   -1  'True
      Top             =   9600
      Width           =   1920
   End
   Begin VB.Label Label18 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00200005&
      Caption         =   "Channel"
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
      Left            =   2040
      TabIndex        =   33
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label Label17 
      BackColor       =   &H00200005&
      Caption         =   "MHz"
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
      Left            =   4920
      TabIndex        =   32
      Top             =   9120
      Width           =   735
   End
   Begin VB.Label Label16 
      BackColor       =   &H00200005&
      Caption         =   "MHz"
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
      Left            =   4920
      TabIndex        =   31
      Top             =   7560
      Width           =   735
   End
   Begin VB.Label Label15 
      BackColor       =   &H00200005&
      Caption         =   "MHz"
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
      Left            =   4920
      TabIndex        =   30
      Top             =   4440
      Width           =   735
   End
   Begin VB.Label Label14 
      BackColor       =   &H00200005&
      Caption         =   "MHz"
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
      Left            =   4920
      TabIndex        =   29
      Top             =   6000
      Width           =   735
   End
   Begin VB.Label Label13 
      BackColor       =   &H00200005&
      Caption         =   "MHz"
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
      Left            =   4920
      TabIndex        =   28
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label12 
      BackColor       =   &H00200005&
      Caption         =   "dBm"
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
      Left            =   10920
      TabIndex        =   27
      Top             =   8280
      Width           =   855
   End
   Begin VB.Label Label11 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00200005&
      Caption         =   "1MHz"
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
      Left            =   6000
      TabIndex        =   25
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label10 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00200005&
      Caption         =   "10MHz Step"
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
      Left            =   7800
      TabIndex        =   24
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label8 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   615
      Left            =   2280
      TabIndex        =   16
      Top             =   9120
      Width           =   375
   End
   Begin VB.Label Label7 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   615
      Left            =   2280
      TabIndex        =   13
      Top             =   7560
      Width           =   375
   End
   Begin VB.Label Label6 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   615
      Left            =   2280
      TabIndex        =   10
      Top             =   6000
      Width           =   375
   End
   Begin VB.Label Label5 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   615
      Left            =   2280
      TabIndex        =   7
      Top             =   4440
      Width           =   375
   End
   Begin VB.Label Label4 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   615
      Left            =   2280
      TabIndex        =   4
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00200005&
      Caption         =   "Signal Generator"
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
      Left            =   -360
      TabIndex        =   2
      Top             =   120
      Width           =   15735
   End
   Begin VB.Label Label2 
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
      Left            =   9600
      TabIndex        =   1
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00200005&
      Caption         =   "Frequency (0~80MHz)"
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
      Left            =   6000
      TabIndex        =   0
      Top             =   1440
      Width           =   3255
   End
End
Attribute VB_Name = "SigGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sigVal As Double
Dim freq As Byte
Dim Amp As Byte
Dim w As Single
Dim t As Single
Dim a As Integer
Dim b As Integer
Dim c As Integer
Dim d As Integer
Dim e As Integer
Dim f As Integer
Dim g As Integer
Dim h As Integer
Dim i As Integer
Dim j As Integer
Dim iSwitchLed1 As Boolean
Dim iSwitchLed2 As Boolean
Dim iSwitchLed3 As Boolean
Dim iSwitchLed4 As Boolean
Dim iSwitchLed5 As Boolean

Dim FreqCheck As Boolean
Dim UseCheck As Boolean
Dim iSwitchLever1 As Boolean

Dim FreqSum1 As Integer
Dim FreqSum2 As Integer
Dim FreqSum3 As Integer
Dim FreqSum4 As Integer
Dim FreqSum5 As Integer

Dim Freqchk1 As Integer
Dim Freqchk2 As Integer
Dim Freqchk3 As Integer
Dim Freqchk4 As Integer
Dim Freqchk5 As Integer

Private Sub Image1_Click()
Open App.Path & "\Setting.txt" For Output As #1
Print #1, a, c, e, g, i, j, h, f, d, b
Write #1, iSwitchLed1, iSwitchLed2, iSwitchLed3, iSwitchLed4, iSwitchLed5, Amp, iSwitchLever1
Close #1
Unload Me
End Sub

Private Sub iSliderX1_OnPositionChangeUser()
Amp = iSliderX1.Position
Call cmd
End Sub

Private Sub iSwitchLedX1_OnChange()
Dim result As Byte
If iSwitchLedX1.Active = False Then
    iSwitchLedX1.Caption = "Start"
    result = IO1.Close
    Timer1.Enabled = False
    iLedSpiralX1.Position = 26
    'iLCDMatrixX1.Text = "STOP"
    'iSwitchRocker3WayX1.Value = 0
    'iSwitchRocker3WayX2.Value = 0
    'iSwitchRocker3WayX3.Value = 0
    'iSwitchRocker3WayX4.Value = 0
    'iSwitchRocker3WayX5.Value = 0
    'iSwitchRocker3WayX7.Value = 0
    'iSwitchRocker3WayX8.Value = 0
    'iSwitchRocker3WayX9.Value = 0
    'iSwitchRocker3WayX10.Value = 0
    'iSwitchRocker3WayX11.Value = 0
    'Freq1txt.Text = 0
    'Freq2txt.Text = 0
    'Freq3txt.Text = 0
    'Freq4txt.Text = 0
    'Freq5txt.Text = 0
    'FreqSum1 = 0
    'FreqSum2 = 0
    'FreqSum3 = 0
    'FreqSum4 = 0
    'FreqSum5 = 0
    'iSliderX1.Position = 0
   'dBtxt.Text = 0
ElseIf iSwitchLedX1.Active = True Then
iSwitchLedX1.Caption = "Stop"
If iSwitchLed1 = False And iSwitchLed2 = False And iSwitchLed3 = False And iSwitchLed4 = False And iSwitchLed5 = False Then
    UseCheck = False
Else
    UseCheck = True
End If
    If UseCheck = True Then
        If FreqCheck = True Then
            iLedSpiralX1.Position = 26
            'iLCDMatrixX1.Text = "RUN"
            result = IO1.Open("COM1:", "baud=115200 parity=N data=8 stop=1")
            Timer1.Enabled = True
            t = 0
            sigVal = 0
        ElseIf FreqCheck = False Then
            MsgBox "주파수가 0또는 중복입니다.", vbExclamation, "Setting"
            iSwitchLedX1.Active = False
        End If
    ElseIf UseCheck = False Then
        MsgBox "사용하는 주파수가 없습니다.", vbExclamation, "Setting"
        iSwitchLedX1.Active = False
    End If
End If
End Sub

Private Sub iSwitchLeverX1_OnChange()
If iSwitchLeverX1.Active Then
    iSwitchLever1 = True
Else
    iSwitchLever1 = False
End If
End Sub

Private Sub iSwitchRocker3WayX1_OnValueChange()
    If iSwitchLed1 = True Then
        a = iSwitchRocker3WayX1.Value
        Call cmd
    ElseIf iSwitchRocker3WayX1.Value <> 0 Then
        iSwitchRocker3WayX1.Value = 0
        MsgBox "사용여부를 확인하세요.", vbExclamation, "Setting"
    End If
End Sub
Private Sub iSwitchRocker3WayX11_OnValueChange()
    If iSwitchLed1 = True Then
        b = iSwitchRocker3WayX11.Value
        Call cmd
    ElseIf iSwitchRocker3WayX11.Value <> 0 Then
        iSwitchRocker3WayX11.Value = 0
        MsgBox "사용여부를 확인하세요.", vbExclamation, "Setting"
        
    End If
End Sub
Private Sub iSwitchRocker3WayX2_OnValueChange()
    If iSwitchLed2 = True Then
        c = iSwitchRocker3WayX2.Value
        Call cmd
    ElseIf iSwitchRocker3WayX2.Value <> 0 Then
        iSwitchRocker3WayX2.Value = 0
        MsgBox "사용여부를 확인하세요.", vbExclamation, "Setting"
        
    End If
End Sub
Private Sub iSwitchRocker3WayX10_OnValueChange()
    If iSwitchLed2 = True Then
        d = iSwitchRocker3WayX10.Value
        Call cmd
    ElseIf iSwitchRocker3WayX10.Value <> 0 Then
        iSwitchRocker3WayX10.Value = 0
        MsgBox "사용여부를 확인하세요.", vbExclamation, "Setting"
        
    End If
End Sub
Private Sub iSwitchRocker3WayX3_OnValueChange()
    If iSwitchLed3 = True Then
        e = iSwitchRocker3WayX3.Value
        Call cmd
    ElseIf iSwitchRocker3WayX3.Value <> 0 Then
        iSwitchRocker3WayX3.Value = 0
        MsgBox "사용여부를 확인하세요.", vbExclamation, "Setting"
        
    End If
End Sub
Private Sub iSwitchRocker3WayX9_OnValueChange()
    If iSwitchLed3 = True Then
        f = iSwitchRocker3WayX9.Value
        Call cmd
    ElseIf iSwitchRocker3WayX9.Value <> 0 Then
        iSwitchRocker3WayX9.Value = 0
        MsgBox "사용여부를 확인하세요.", vbExclamation, "Setting"
        
    End If
End Sub
Private Sub iSwitchRocker3WayX4_OnValueChange()
    If iSwitchLed4 = True Then
        g = iSwitchRocker3WayX4.Value
        Call cmd
    ElseIf iSwitchRocker3WayX4.Value <> 0 Then
         iSwitchRocker3WayX4.Value = 0
         MsgBox "사용여부를 확인하세요.", vbExclamation, "Setting"
       
    End If
End Sub
Private Sub iSwitchRocker3WayX8_OnValueChange()
    If iSwitchLed4 = True Then
        h = iSwitchRocker3WayX8.Value
        Call cmd
    ElseIf iSwitchRocker3WayX8.Value <> 0 Then
        iSwitchRocker3WayX8.Value = 0
        MsgBox "사용여부를 확인하세요.", vbExclamation, "Setting"
        
    End If
End Sub
Private Sub iSwitchRocker3WayX5_OnValueChange()
    If iSwitchLed5 = True Then
        i = iSwitchRocker3WayX5.Value
        Call cmd
    ElseIf iSwitchRocker3WayX5.Value <> 0 Then
        iSwitchRocker3WayX5.Value = 0
        MsgBox "사용여부를 확인하세요.", vbExclamation, "Setting"
        
    End If
End Sub
Private Sub iSwitchRocker3WayX7_OnValueChange()
    If iSwitchLed5 = True Then
        j = iSwitchRocker3WayX7.Value
        Call cmd
    ElseIf iSwitchRocker3WayX7.Value <> 0 Then
        iSwitchRocker3WayX7.Value = 0
        MsgBox "사용여부를 확인하세요.", vbExclamation, "Setting"
        
    End If
End Sub
Private Sub cmd()

FreqSum1 = a + b
FreqSum2 = c + d
FreqSum3 = e + f
FreqSum4 = g + h
FreqSum5 = i + j
    If iSwitchLed1 = True And FreqSum1 >= 0 Then
        Freqchk1 = FreqSum1
    Else
        Freqchk1 = 100
    End If
    If iSwitchLed2 = True And FreqSum2 >= 0 Then
        Freqchk2 = FreqSum2
    Else
        Freqchk2 = 101
    End If
    If iSwitchLed3 = True And FreqSum3 >= 0 Then
        Freqchk3 = FreqSum3
    Else
        Freqchk3 = 103
    End If
    If iSwitchLed4 = True And FreqSum4 >= 0 Then
        Freqchk4 = FreqSum4
    Else
        Freqchk4 = 104
    End If
    If iSwitchLed5 = True And FreqSum5 >= 0 Then
        Freqchk5 = FreqSum5
    Else: Freqchk5 = 105
    End If
    
    If (Freqchk1 = Freqchk2) Or (Freqchk1 = Freqchk3) Or (Freqchk1 = Freqchk4) Or (Freqchk1 = Freqchk5) Or (Freqchk2 = Freqchk3) Or (Freqchk2 = Freqchk4) Or (Freqchk2 = Freqchk5) Or (Freqchk3 = Freqchk4) Or (Freqchk3 = Freqchk5) Or (Freqchk4 = Freqchk5) Then
        FreqCheck = False
    ElseIf Freqchk1 = 0 Or Freqchk2 = 0 Or Freqchk3 = 0 Or Freqchk4 = 0 Or Freqchk5 = 0 Then
        FreqCheck = False
    Else
        FreqCheck = True
    End If



    If FreqSum1 >= 0 And FreqSum1 <= 80 And iSwitchLed1 = True Then
        Freq1txt.Text = Freqchk1
    ElseIf FreqSum1 < 0 Then
        MsgBox "범위를 벗어납니다.", vbExclamation, "Setting"
        iSwitchRocker3WayX1.Value = 0
        iSwitchRocker3WayX11.Value = 0
    ElseIf FreqSum1 > 80 Then
        MsgBox "범위를 벗어납니다.", vbExclamation, "Setting"
        iSwitchRocker3WayX1.Value = 80
        iSwitchRocker3WayX11.Value = 0
    End If


    If FreqSum2 >= 0 And FreqSum2 <= 80 And iSwitchLed2 = True Then
        Freq2txt.Text = Freqchk2
    ElseIf FreqSum2 < 0 Then
        MsgBox "범위를 벗어납니다.", vbExclamation, "Setting"
        iSwitchRocker3WayX2.Value = 0
        iSwitchRocker3WayX10.Value = 0
    ElseIf FreqSum2 > 80 Then
        MsgBox "범위를 벗어납니다.", vbExclamation, "Setting"
        iSwitchRocker3WayX2.Value = 80
        iSwitchRocker3WayX10.Value = 0
    End If

    If FreqSum3 >= 0 And FreqSum3 <= 80 And iSwitchLed3 = True Then
        Freq3txt.Text = Freqchk3
    ElseIf FreqSum3 < 0 Then
        MsgBox "범위를 벗어납니다.", vbExclamation, "Setting"
        iSwitchRocker3WayX3.Value = 0
        iSwitchRocker3WayX9.Value = 0
    ElseIf FreqSum3 > 80 Then
        MsgBox "범위를 벗어납니다.", vbExclamation, "Setting"
        iSwitchRocker3WayX3.Value = 80
        iSwitchRocker3WayX9.Value = 0
    End If

    If FreqSum4 >= 0 And FreqSum4 <= 80 And iSwitchLed4 = True Then
        Freq4txt.Text = Freqchk4
    ElseIf FreqSum4 < 0 Then
        MsgBox "범위를 벗어납니다.", vbExclamation, "Setting"
        iSwitchRocker3WayX4.Value = 0
        iSwitchRocker3WayX8.Value = 0
    ElseIf FreqSum4 > 80 Then
        MsgBox "범위를 벗어납니다.", vbExclamation, "Setting"
        iSwitchRocker3WayX4.Value = 80
        iSwitchRocker3WayX8.Value = 0
    End If

    If FreqSum5 >= 0 And FreqSum5 <= 80 And iSwitchLed5 = True Then
        Freq5txt.Text = Freqchk5
    ElseIf FreqSum5 < 0 Then
        MsgBox "범위를 벗어납니다.", vbExclamation, "Setting"
        iSwitchRocker3WayX5.Value = 0
        iSwitchRocker3WayX7.Value = 0
    ElseIf FreqSum5 > 80 Then
        MsgBox "범위를 벗어납니다.", vbExclamation, "Setting"
        iSwitchRocker3WayX5.Value = 80
        iSwitchRocker3WayX7.Value = 0
    End If
    dBtxt.Text = Amp
End Sub
Private Sub Form_Load()
Timer1.Enabled = False
sigVal = 0
Open App.Path & "\Setting.txt" For Input As #1
Input #1, a, c, e, g, i, j, h, f, d, b
Input #1, iSwitchLed1, iSwitchLed2, iSwitchLed3, iSwitchLed4, iSwitchLed5, Amp, iSwitchLever1

If iSwitchLed1 = True Then
    iSwitchLedX2.Active = True
Else: iSwitchLedX2.Active = False
End If
If iSwitchLed2 = True Then
    iSwitchLedX3.Active = True
Else: iSwitchLedX3.Active = False
End If
If iSwitchLed3 = True Then
    iSwitchLedX4.Active = True
Else: iSwitchLedX4.Active = False
End If
If iSwitchLed4 = True Then
    iSwitchLedX5.Active = True
Else: iSwitchLedX5.Active = False
End If
If iSwitchLed5 = True Then
    iSwitchLedX6.Active = True
Else: iSwitchLedX6.Active = False
End If
If iSwitchLever1 = True Then
    iSwitchLeverX1.Active = True
Else: iSwitchLeverX1.Active = False
End If

iSwitchRocker3WayX1.Value = a
iSwitchRocker3WayX2.Value = c
iSwitchRocker3WayX3.Value = e
iSwitchRocker3WayX4.Value = g
iSwitchRocker3WayX5.Value = i
iSwitchRocker3WayX7.Value = j
iSwitchRocker3WayX8.Value = h
iSwitchRocker3WayX9.Value = f
iSwitchRocker3WayX10.Value = d
iSwitchRocker3WayX11.Value = b
iSliderX1.Position = Amp

If iSwitchLedX2.Active = False And iSwitchLedX3.Active = False And iSwitchLedX4.Active = False And iSwitchLedX5.Active = False And iSwitchLedX6.Active = False Then
    iSwitchLedX2.Active = True
    iSwitchRocker3WayX11.Value = 1
End If

Close #1
End Sub

Private Sub iSwitchLedX2_OnChange()
If iSwitchLedX2.Active Then
    iSwitchLed1 = True
Else
    iSwitchLed1 = False
End If
Call cmd
End Sub

Private Sub iSwitchLedX3_OnChange()
If iSwitchLedX3.Active Then
    iSwitchLed2 = True
Else
    iSwitchLed2 = False
End If
Call cmd
End Sub
Private Sub iSwitchLedX4_OnChange()
If iSwitchLedX4.Active Then
    iSwitchLed3 = True
Else
    iSwitchLed3 = False
End If
Call cmd
End Sub
Private Sub iSwitchLedX5_OnChange()
If iSwitchLedX5.Active Then
    iSwitchLed4 = True
Else
    iSwitchLed4 = False
End If
Call cmd
End Sub
Private Sub iSwitchLedX6_OnChange()
If iSwitchLedX6.Active Then
    iSwitchLed5 = True
Else
    iSwitchLed5 = False
End If
Call cmd
End Sub
Private Sub Timer1_Timer()
Dim m As Double
Dim pi As Single
    If iLedSpiralX1.Position <> 100 Then
        iLedSpiralX1.Position = iLedSpiralX1.Position + 1
    ElseIf iLedSpiralX1.Position = 100 Then
        iLedSpiralX1.Position = 0
    End If
    pi = 3.14159
    m = 2 * pi * freq * t 'scal factor 10
    t = t + (Timer1.Interval / 1000)
    sigVal = Amp * 10 * Sin(m)
    IO1.WriteByte (sigVal)
    'Timer1.Enabled = True
End Sub

