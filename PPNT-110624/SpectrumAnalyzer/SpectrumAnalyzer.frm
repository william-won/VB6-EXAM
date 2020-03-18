VERSION 5.00
Object = "{D1120C7B-28C7-11D3-85BF-00105AC8B715}#1.0#0"; "iStripChartXControl.ocx"
Object = "{A8B345A0-74B5-11D3-85C2-00105AC8B715}#1.0#0"; "iProfessionalLibrary.ocx"
Object = "{0A362340-2E5E-11D3-85BF-00105AC8B715}#1.0#0"; "isDigitalLibrary.ocx"
Begin VB.Form frmMainForm 
   BackColor       =   &H00200005&
   BorderStyle     =   0  '없음
   Caption         =   "Spectrum Analyzer"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   FillColor       =   &H00FFFFFF&
   FillStyle       =   0  '단색
   LinkTopic       =   "Form1"
   Picture         =   "SpectrumAnalyzer.frx":0000
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text3 
      BackColor       =   &H00200005&
      BorderStyle     =   0  '없음
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   "dBm"
      Top             =   8520
      Width           =   615
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8400
      Style           =   1  '그래픽
      TabIndex        =   9
      Top             =   9960
      Width           =   1935
   End
   Begin VB.CommandButton cmdRst 
      BackColor       =   &H00FFC0C0&
      Caption         =   "reset Max Hold"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6360
      Style           =   1  '그래픽
      TabIndex        =   8
      Top             =   9960
      Width           =   1935
   End
   Begin VB.CommandButton CmdSetting 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Setting"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4320
      Style           =   1  '그래픽
      TabIndex        =   3
      Top             =   9960
      Width           =   1935
   End
   Begin VB.Timer Timer 
      Interval        =   5
      Left            =   14880
      Top             =   600
   End
   Begin VB.CommandButton cmdShowHideToolbar 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Show Grid"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2280
      Style           =   1  '그래픽
      TabIndex        =   0
      Top             =   9960
      Width           =   1935
   End
   Begin iStripChartXControl.iStripChartX iStripChartX1 
      Height          =   8655
      Left            =   2280
      TabIndex        =   1
      Top             =   1080
      Width           =   12975
      AxisGridColor   =   16777215
      TitleText       =   ""
      TitleMargin     =   0
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      XAxisMin        =   0
      XAxisMax        =   80
      XAxisMargin     =   1
      XAxisLabelMargin=   5
      BeginProperty XAxisLabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      XAxisLabelPrecision=   1
      XAxisTitle      =   "Frequency (MHz)"
      XAxisTitleMargin=   0
      BeginProperty XAxisTitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      XAxisTickMajorCount=   16
      XAxisTickMajorLength=   7
      XAxisTickMajorColor=   16777215
      XAxisTickMinorCount=   4
      XAxisTickMinorLength=   3
      XAxisTickMinorColor=   16777215
      YAxisMin        =   -60
      YAxisMax        =   20
      YAxisMargin     =   5
      YAxisLabelMargin=   5
      BeginProperty YAxisLabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      YAxisLabelPrecision=   1
      YAxisTitle      =   "Amplitude (dBm)"
      YAxisTitleMargin=   0
      BeginProperty YAxisTitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      YAxisTickMajorCount=   6
      YAxisTickMajorLength=   7
      YAxisTickMajorColor=   16777215
      YAxisTickMinorCount=   4
      YAxisTickMinorLength=   3
      YAxisTickMinorColor=   16777215
      GridLineStyle   =   2
      OuterMarginLeft =   10
      OuterMarginTop  =   10
      OuterMarginRight=   10
      OuterMarginBottom=   10
      LegendWidth     =   80
      LegendMargin    =   10
      BeginProperty LegendFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowGrid        =   -1  'True
      ShowLegend      =   0   'False
      ShowToolBar     =   -1  'True
      AutoScrollEnabled=   -1  'True
      AutoScrollType  =   0
      AutoScrollStepSize=   0
      AutoScaleEnabled=   -1  'True
      AutoScaleHysterisis=   0
      BeginProperty ToolBarActiveModeFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty ToolBarInactiveModeFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ToolBarMode     =   4
      CursorChannel   =   0
      CursorColor     =   65535
      CursorChannelBackgroundColor=   65535
      CursorChannelFontColor=   0
      BorderStyle     =   2
      BackGroundColor =   0
      TitleFontColor  =   16777215
      XAxisTitleFontColor=   16777215
      YAxisTitleFontColor=   16777215
      XAxisLabelFontColor=   16777215
      YAxisLabelFontColor=   16777215
      LegendFontColor =   16777215
      ToolbarActiveModeFontColor=   65535
      ToolbarInactiveModeFontColor=   16776960
      XAxisDateTimeEnabled=   0   'False
      XAxisDateTimeFormatString=   "hh:nn:ssam/pm"
      GridBackGroundColor=   0
      PrinterOrientation=   0
      PrinterMarginLeft=   0.5
      PrinterMarginTop=   0.5
      PrinterMarginRight=   0.5
      PrinterMarginBottom=   0.5
      RestoreXYAxisOnPlotMode=   0   'False
      MaxBufferSize   =   0
      MinBufferSize   =   0
      CursorIndex     =   0
      PrinterCommentLineSpacing=   0
      BeginProperty PrinterCommentLinesFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      EnableDataDrawMinMax=   -1  'True
      CursorHideAllOtherChannels=   0   'False
      XAxisShow       =   -1  'True
      AutoScaleMinAdjustEnabled=   0   'False
      AutoScaleMaxAdjustEnabled=   0   'False
      DiscontinuousDataEnabled=   -1  'True
      AutoScrollFirstStyle=   1
      PrecisionStyle  =   0
      InterpolateMissingDataPoints=   0   'False
      PrinterShowDialog=   -1  'True
      YAxisShow       =   -1  'True
      YAxisLabelWidth =   20
      YAxisLabelWidthFixed=   0   'False
      YAxisReverseScale=   0   'False
      UpdateFrameRate =   60
      Object.Visible         =   -1  'True
      ElapsedStartTime=   40679.9046099074
      PrinterCommentLinesFontColor=   -16777208
      OptionSaveAllProperties=   0   'False
      AddYChannel1Now =   0
      AutoFrameRate   =   0   'False
      Enabled         =   -1  'True
      Object.Width           =   865
      Object.Height          =   577
      ChannelCount    =   2
      Channel0Title   =   "min"
      Channel0Color   =   65280
      Channel0LineStyle=   0
      Channel0LineWidth=   1
      Channel1Title   =   "max"
      Channel1Color   =   4227327
      Channel1LineStyle=   0
      Channel1LineWidth=   1
      TitleFontName   =   "MS Sans Serif"
      TitleFontSize   =   8
      TitleFontBold   =   0   'False
      TitleFontItalic =   0   'False
      TitleFontUnderline=   0   'False
      TitleFontStrikeOut=   0   'False
      XAxisTitleFontName=   "Arial"
      XAxisTitleFontSize=   12
      XAxisTitleFontBold=   -1  'True
      XAxisTitleFontItalic=   0   'False
      XAxisTitleFontUnderline=   0   'False
      XAxisTitleFontStrikeOut=   0   'False
      YAxisTitleFontName=   "Arial"
      YAxisTitleFontSize=   12
      YAxisTitleFontBold=   -1  'True
      YAxisTitleFontItalic=   0   'False
      YAxisTitleFontUnderline=   0   'False
      YAxisTitleFontStrikeOut=   0   'False
      XAxisLabelFontName=   "Arial"
      XAxisLabelFontSize=   10
      XAxisLabelFontBold=   0   'False
      XAxisLabelFontItalic=   0   'False
      XAxisLabelFontUnderline=   0   'False
      XAxisLabelFontStrikeOut=   0   'False
      YAxisLabelFontName=   "Arial"
      YAxisLabelFontSize=   10
      YAxisLabelFontBold=   0   'False
      YAxisLabelFontItalic=   0   'False
      YAxisLabelFontUnderline=   0   'False
      YAxisLabelFontStrikeOut=   0   'False
      LegendFontName  =   "MS Sans Serif"
      LegendFontSize  =   8
      LegendFontBold  =   0   'False
      LegendFontItalic=   0   'False
      LegendFontUnderline=   0   'False
      LegendFontStrikeOut=   0   'False
      ToolbarActiveModeFontName=   "MS Sans Serif"
      ToolbarActiveModeFontSize=   8
      ToolbarActiveModeFontBold=   -1  'True
      ToolbarActiveModeFontItalic=   0   'False
      ToolbarActiveModeFontUnderline=   0   'False
      ToolbarActiveModeFontStrikeOut=   0   'False
      ToolbarInactiveModeFontName=   "MS Sans Serif"
      ToolbarInactiveModeFontSize=   8
      ToolbarInactiveModeFontBold=   -1  'True
      ToolbarInactiveModeFontItalic=   0   'False
      ToolbarInactiveModeFontUnderline=   0   'False
      ToolbarInactiveModeFontStrikeOut=   0   'False
   End
   Begin iProfessionalLibrary.iSwitchLeverX iSwitchLeverX1 
      Height          =   2295
      Left            =   720
      TabIndex        =   4
      Top             =   4560
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
      Left            =   120
      TabIndex        =   10
      Top             =   1080
      Width           =   1935
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
      Glyph           =   "SpectrumAnalyzer.frx":5355
      BorderSize      =   5
      BorderHighlightColor=   16761024
      BorderShadowColor=   16761024
      BackGroundColor =   8421376
      OptionSaveAllProperties=   0   'False
      AutoFrameRate   =   0   'False
      Object.Width           =   129
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
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  '없음
      Height          =   9735
      Left            =   0
      Picture         =   "SpectrumAnalyzer.frx":53AB
      ScaleHeight     =   9735
      ScaleWidth      =   15360
      TabIndex        =   11
      Top             =   0
      Width           =   15355
      Begin VB.TextBox Text2 
         BackColor       =   &H00200005&
         BorderStyle     =   0  '없음
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "MHz"
         Top             =   8040
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00200005&
         BorderStyle     =   0  '없음
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   990
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   12
         Text            =   "SpectrumAnalyzer.frx":D646
         Top             =   7920
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "Spectrum Analyzer"
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
      TabIndex        =   2
      Top             =   360
      Width           =   15135
   End
   Begin VB.Image ImgExit 
      Height          =   1950
      Left            =   13440
      Picture         =   "SpectrumAnalyzer.frx":D663
      Stretch         =   -1  'True
      Top             =   9600
      Width           =   1920
   End
   Begin VB.Label Label40 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00200005&
      Caption         =   "Input Port"
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
      Left            =   240
      TabIndex        =   7
      Top             =   3480
      Width           =   1815
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
      TabIndex        =   6
      Top             =   3960
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
      Left            =   240
      TabIndex        =   5
      Top             =   7080
      Width           =   1695
   End
End
Attribute VB_Name = "frmMainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*******************************************************
'
'       iExample0002- Strip Chart Example
'
'       Copyright (c) 2000 Iocomp Software
'
'*******************************************************

'This simple example illustrates using the Strip Chart Control.

'This example also demonstrates the following...
'1) Adding Channels
'2) Modifying basic control properties
'3) Adding Data to Strip Chart
'4) Clearing Data from Strip Chart
'5) Showing/Hiding toolbar

'Data is added to the strip chart from randomly generated data.

'REMEMBER: Integers(32-bit) specified in the help files refers to the Visual Basic
'          Long type. (Integer in VB is a 16-bit integer, Long is 32-bit integer)

'Permission is granted to distribute this source code without restriction.
'http://www.iocomp.com

'==================
' Global Variables
'==================
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Ymax As Long
Private Xindex As Long
Private StartGrid As Boolean

Private Const MAXLEN = 801 'iStripChartX1.XAxisMax*10+1
Private CurrentTimeCounter As Double
Private MinMaxData(MAXLEN) As Long
Private CurrentIndexTime As Long
Private Sub cmdRst_Click()
    iStripChartX1.ClearData
    CurrentTimeCounter = 0
    CurrentIndexTime = 0
End Sub

Private Sub cmdSave_Click()
Timer.Enabled = True

'iStripChartX1.PrintChart
Picture1.AutoRedraw = True
 
    'iStripChartX1.SaveImageToBitmap ("Hello1.bmp")
BitBlt Picture1.hDC, 0, 0, Screen.Width / Screen.TwipsPerPixelX, Screen.Height / Screen.TwipsPerPixelY, GetDC(GetDesktopWindow), 0, 0, &HCC0020
SavePicture Picture1.Image, Format(Date, "yyyy-mm-dd") & Format(Time, "_hh시mm분ss초") & ".bmp"
End Sub
Private Sub CmdSetting_Click()
    iSwitchLedX1.Active = False
    iSwitchLedX1.Caption = "Start"
    StartGrid = False
    frmSetting.Show
End Sub

'=====================================================================================
' Name: cmdShowHideToolbar_Click
' Description: Handles click event on Show/Hide Toolbar button
'=====================================================================================
Private Sub cmdShowHideToolbar_Click()
    If iStripChartX1.ShowToolBar = True Then
        iStripChartX1.ShowToolBar = False
    Else
        iStripChartX1.ShowToolBar = True
        iStripChartX1.ToolBarMode = iscmCursor
    End If
End Sub

'=====================================================================================
' Name: Form_Load
' Description: Handles loading event for main form.  Sets initial values of controls
'=====================================================================================
Private Sub Form_Load()
    'Set the initial properties of controls on the page
    'Note: you can also set all of these properties and add channels at design-time using the built-in property editors.
    '      we have set all of the properties at run-time for example purposes only so that you can see all of the modifications
    '      made to the controls from their defaults.
    
    'Add two channels
    'iStripChartX1.AddChannel "Channel 1", vbRed, iclsDash, 1 'red solid line
    'iStripChartX1.AddChannel "Channel 2", vbBlue, iclsDash, 1 'blue solid line
    'iStripChartX1.AddChannel "Min-Max", vbYellow, iclsSolid, 1 'yellow dashed line
    'Set Legend Width so that the labels are not clipped
    StartGrid = False
    
    iStripChartX1.LegendWidth = 100
    
    'Setup the x and y axis settings
    iStripChartX1.XAxisMax = 80
    iStripChartX1.XAxisMin = 0
    'iStripChartX1.XAxisTitle = "MHz"
    

    iStripChartX1.YAxisMax = 20
    iStripChartX1.YAxisMin = -60
    'iStripChartX1.YAxisTitle = "dBm"
    
    'Set Other Properties
    iStripChartX1.EnableDataDrawMinMax = True
    iStripChartX1.ShowToolBar = False
    
    'iStripChartX1.PrinterShowDialog = True  'Print ?
    Text1.Text = ""
    Dim i As Integer
    For i = 0 To MAXLEN
        MinMaxData(i) = iStripChartX1.YAxisMin
    Next
    
End Sub
Private Sub ImgExit_Click()
Unload Me
Unload frmSetting
End Sub

Private Sub iStripChartX1_OnCursorIndexChange()
        Ymax = MinMaxData(iStripChartX1.CursorIndex)
        Xindex = iStripChartX1.CursorIndex / 10
        Text1.Text = "X = " & Xindex & vbCrLf & "Y = " & Ymax
End Sub

Private Sub iSwitchLedX1_OnChange()
If iSwitchLedX1.Active = True Then
    iSwitchLedX1.Caption = "Stop"
    StartGrid = True
ElseIf iSwitchLedX1.Active = False Then
    iSwitchLedX1.Caption = "Start"
    StartGrid = False
End If
End Sub

'=====================================================================================
' Name: Timer_Timer
' Description: Handles timer event.  Timer should call sub every 100ms
'=====================================================================================
Private Sub Timer_Timer()
    Dim RandomData1 As Double
    'Dim RandomData2 As Long
If StartGrid = True Then
    
    'Generate Random Data
    '====================================================================
    'IGNORE THIS CODE.  YOU WON'T NEED IT IN YOUR PROGRAM
    'THIS IS FOR GENERATING RANDOM DATA ONLY
    RandomData1 = Rnd() * 50 - 50
    'RandomData2 = Rnd() * 500 - 250
    'RandomData2 = Sin(CurrentTimeCounter) * 250
    '====================================================================
    'END GENERATE RANDOM DATA

    iStripChartX1.BeginUpdate 'Stops painting to the Strip Chart channel area
    
    If iStripChartX1.IndexCount <= iStripChartX1.XAxisMax * 10 Then
    CurrentIndexTime = iStripChartX1.AddIndexTime(CurrentTimeCounter)  'Add an index time.  When we add data
                                                                       'to the chart, all of the channel
                                                                       'data points are synchronized to this
                                                                       'specified time-point.
    Else
        CurrentIndexTime = CurrentTimeCounter * 10
    End If
    'Rember that Channel numbers are 0 based, so...
    iStripChartX1.SetChannelData 0, CurrentIndexTime, RandomData1 'Add data to channel 1
    If RandomData1 > MinMaxData(CurrentIndexTime) Then 'min-max holder
        MinMaxData(CurrentIndexTime) = RandomData1
    End If
    
    iStripChartX1.SetChannelData 1, CurrentIndexTime, (MinMaxData(CurrentIndexTime) + 7) 'Add data to channel 2

    CurrentTimeCounter = CurrentTimeCounter + 0.1 'Increment the global time variable (seconds)
    If CurrentTimeCounter > (iStripChartX1.XAxisMax) Then
        CurrentTimeCounter = 0
        CurrentIndexTime = 0
        'iStripChartX1.ClearData 'kate
        
    End If
    
    
   
    iStripChartX1.EndUpdate 'Resumes painting to the Strip Chart channel area
End If
End Sub










