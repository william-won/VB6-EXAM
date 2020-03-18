VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Setting"
   ClientHeight    =   6315
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14760
   LinkTopic       =   "Form2"
   ScaleHeight     =   6315
   ScaleWidth      =   14760
   StartUpPosition =   3  'Windows ±âº»°ª
   Begin VB.TextBox txtMACAddress 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   79
      Text            =   "00:00:00:00:00:00"
      Top             =   240
      Width           =   2055
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   9000
      TabIndex        =   78
      Top             =   5400
      Width           =   2175
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   495
      Left            =   6360
      TabIndex        =   77
      Top             =   5400
      Width           =   2175
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "Change"
      Height          =   495
      Left            =   3480
      TabIndex        =   76
      Top             =   5400
      Width           =   2295
   End
   Begin VB.Frame Frame13 
      Caption         =   "Parity bits"
      Height          =   1455
      Left            =   11640
      TabIndex        =   72
      Top             =   3480
      Width           =   3015
      Begin VB.CheckBox Check11 
         Caption         =   "Fast Recovery Flag"
         Height          =   255
         Left            =   120
         TabIndex        =   75
         Top             =   960
         Width           =   2175
      End
      Begin VB.CheckBox Check10 
         Caption         =   "New Cell Request Flag"
         Height          =   255
         Left            =   120
         TabIndex        =   74
         Top             =   600
         Width           =   2295
      End
      Begin VB.CheckBox Check9 
         Caption         =   "NEP Flag"
         Height          =   255
         Left            =   120
         TabIndex        =   73
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame12 
      Caption         =   "Mode"
      Height          =   1455
      Left            =   10320
      TabIndex        =   70
      Top             =   3480
      Width           =   1095
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   240
         TabIndex        =   71
         Top             =   600
         Width           =   615
      End
   End
   Begin VB.Frame Frame11 
      Caption         =   "MAC Type"
      Height          =   1455
      Left            =   7440
      TabIndex        =   66
      Top             =   3480
      Width           =   2655
      Begin VB.OptionButton Option34 
         Caption         =   "Access MAC"
         Height          =   255
         Left            =   240
         TabIndex        =   69
         Top             =   600
         Width           =   1575
      End
      Begin VB.OptionButton Option33 
         Caption         =   "In-Home MAC"
         Height          =   255
         Left            =   240
         TabIndex        =   68
         Top             =   960
         Width           =   1575
      End
      Begin VB.OptionButton Option30 
         Caption         =   "Basic MAC(CMAC)"
         Height          =   255
         Left            =   240
         TabIndex        =   67
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "UART Configuration"
      Height          =   3015
      Left            =   7440
      TabIndex        =   38
      Top             =   240
      Width           =   7215
      Begin VB.Frame Frame10 
         Caption         =   "Serial Data PLC Tx Type"
         Height          =   975
         Left            =   3720
         TabIndex        =   62
         Top             =   1920
         Width           =   3015
         Begin VB.OptionButton Option29 
            Caption         =   "Repeater"
            Height          =   180
            Left            =   1680
            TabIndex        =   65
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton Option22 
            Caption         =   "Meter Gateway"
            Height          =   180
            Left            =   240
            TabIndex        =   64
            Top             =   600
            Width           =   2055
         End
         Begin VB.OptionButton Option21 
            Caption         =   "Default"
            Height          =   180
            Left            =   240
            TabIndex        =   63
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Parity bits"
         Height          =   975
         Left            =   240
         TabIndex        =   56
         Top             =   1920
         Width           =   3015
         Begin VB.CheckBox Check8 
            Caption         =   "Parity Enable"
            Height          =   495
            Left            =   120
            TabIndex        =   61
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton Option32 
            Caption         =   "Odd"
            Height          =   180
            Left            =   1200
            TabIndex        =   60
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton Option31 
            Caption         =   "Even"
            Height          =   180
            Left            =   1200
            TabIndex        =   59
            Top             =   600
            Width           =   735
         End
         Begin VB.OptionButton Option28 
            Caption         =   "1"
            Height          =   180
            Left            =   2280
            TabIndex        =   58
            Top             =   240
            Width           =   495
         End
         Begin VB.OptionButton Option27 
            Caption         =   "0"
            Height          =   180
            Left            =   2280
            TabIndex        =   57
            Top             =   600
            Width           =   495
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Stop bits"
         Height          =   1575
         Left            =   5400
         TabIndex        =   53
         Top             =   240
         Width           =   1215
         Begin VB.OptionButton Option20 
            Caption         =   "2 bits"
            Height          =   180
            Left            =   120
            TabIndex        =   55
            Top             =   600
            Width           =   975
         End
         Begin VB.OptionButton Option19 
            Caption         =   "1 bits"
            Height          =   180
            Left            =   120
            TabIndex        =   54
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Data bits"
         Height          =   1575
         Left            =   3720
         TabIndex        =   48
         Top             =   240
         Width           =   1215
         Begin VB.OptionButton Option26 
            Caption         =   "5 bits"
            Height          =   180
            Left            =   120
            TabIndex        =   52
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton Option25 
            Caption         =   "6 bits"
            Height          =   180
            Left            =   120
            TabIndex        =   51
            Top             =   600
            Width           =   855
         End
         Begin VB.OptionButton Option24 
            Caption         =   "7 bits"
            Height          =   180
            Left            =   120
            TabIndex        =   50
            Top             =   960
            Width           =   855
         End
         Begin VB.OptionButton Option23 
            Caption         =   "8 bits"
            Height          =   180
            Left            =   120
            TabIndex        =   49
            Top             =   1320
            Width           =   855
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Baud rate(bps)"
         Height          =   1575
         Left            =   240
         TabIndex        =   39
         Top             =   240
         Width           =   3015
         Begin VB.OptionButton Option18 
            Caption         =   "115200"
            Height          =   180
            Left            =   1800
            TabIndex        =   47
            Top             =   1320
            Width           =   975
         End
         Begin VB.OptionButton Option17 
            Caption         =   "38400"
            Height          =   180
            Left            =   1800
            TabIndex        =   46
            Top             =   960
            Width           =   855
         End
         Begin VB.OptionButton Option16 
            Caption         =   "9600"
            Height          =   180
            Left            =   1800
            TabIndex        =   45
            Top             =   600
            Width           =   735
         End
         Begin VB.OptionButton Option15 
            Caption         =   "2400"
            Height          =   180
            Left            =   1800
            TabIndex        =   44
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton Option14 
            Caption         =   "57600"
            Height          =   180
            Left            =   120
            TabIndex        =   43
            Top             =   1320
            Width           =   855
         End
         Begin VB.OptionButton Option13 
            Caption         =   "19200"
            Height          =   180
            Left            =   120
            TabIndex        =   42
            Top             =   960
            Width           =   855
         End
         Begin VB.OptionButton Option12 
            Caption         =   "4800"
            Height          =   180
            Left            =   120
            TabIndex        =   41
            Top             =   600
            Width           =   855
         End
         Begin VB.OptionButton Option11 
            Caption         =   "1200"
            Height          =   180
            Left            =   120
            TabIndex        =   40
            Top             =   240
            Width           =   735
         End
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "General Configuration"
      Height          =   3015
      Left            =   120
      TabIndex        =   14
      Top             =   2280
      Width           =   7215
      Begin VB.TextBox Text6 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6360
         TabIndex        =   36
         Text            =   "0"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4440
         TabIndex        =   34
         Text            =   "0"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   5040
         TabIndex        =   30
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   6000
         TabIndex        =   29
         Top             =   720
         Width           =   375
      End
      Begin VB.Frame Frame6 
         Caption         =   "Bank Indicator"
         Height          =   615
         Left            =   3600
         TabIndex        =   26
         Top             =   1200
         Width           =   3495
         Begin VB.OptionButton Option10 
            Caption         =   "Bank 0"
            Height          =   255
            Left            =   240
            TabIndex        =   28
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton Option9 
            Caption         =   "Bank 1"
            Height          =   255
            Left            =   1800
            TabIndex        =   27
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Station Mode"
         Height          =   975
         Left            =   3600
         TabIndex        =   21
         Top             =   1920
         Width           =   3495
         Begin VB.OptionButton Option4 
            Caption         =   "Active"
            Height          =   255
            Left            =   240
            TabIndex        =   25
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton Option6 
            Caption         =   "Suspend"
            Height          =   255
            Left            =   1800
            TabIndex        =   24
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton Option7 
            Caption         =   "Factory"
            Height          =   255
            Left            =   240
            TabIndex        =   23
            Top             =   600
            Width           =   1095
         End
         Begin VB.OptionButton Option8 
            Caption         =   "Programming"
            Height          =   255
            Left            =   1800
            TabIndex        =   22
            Top             =   600
            Width           =   1455
         End
      End
      Begin VB.CheckBox Check7 
         Caption         =   "Secondary Interface Enable"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   3135
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Link Restriction Enable"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   720
         Width           =   3135
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Secondary Link Restriction Enable"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1200
         Width           =   3255
      End
      Begin VB.CheckBox Check4 
         Caption         =   "RTS/CTS Enable"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1680
         Width           =   2415
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Tx Notification Enable"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   2160
         Width           =   2655
      End
      Begin VB.CheckBox Check6 
         Caption         =   "Key Transmission Enable"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   2640
         Width           =   2895
      End
      Begin VB.Label Label10 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         Caption         =   "MSDU"
         Height          =   255
         Left            =   5280
         TabIndex        =   37
         Top             =   270
         Width           =   975
      End
      Begin VB.Label Label9 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         Caption         =   "TTL"
         Height          =   255
         Left            =   3600
         TabIndex        =   35
         Top             =   270
         Width           =   615
      End
      Begin VB.Label Label18 
         Caption         =   "Reset Period"
         Height          =   255
         Left            =   3720
         TabIndex        =   33
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label19 
         Caption         =   "Hour"
         Height          =   255
         Left            =   5520
         TabIndex        =   32
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label20 
         Caption         =   "Min"
         Height          =   255
         Left            =   6480
         TabIndex        =   31
         Top             =   840
         Width           =   495
      End
   End
   Begin VB.TextBox Text11 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6600
      TabIndex        =   12
      Text            =   "0"
      Top             =   1290
      Width           =   495
   End
   Begin VB.CheckBox Check1 
      Caption         =   "EU Flag"
      Height          =   375
      Left            =   3840
      TabIndex        =   11
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Device Type"
      Height          =   975
      Left            =   3720
      TabIndex        =   6
      Top             =   240
      Width           =   3495
      Begin VB.OptionButton Option3 
         Caption         =   "MGW"
         Height          =   180
         Left            =   1800
         TabIndex        =   10
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Slave"
         Height          =   180
         Left            =   240
         TabIndex        =   9
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Master"
         Height          =   180
         Left            =   1800
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Default"
         Height          =   180
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.TextBox txtGIDwrite 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Text            =   "00:00:00:00:00:00"
      Top             =   1200
      Width           =   2055
   End
   Begin VB.TextBox txtKEYwrite 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Text            =   "00:00:00:00:00:00"
      Top             =   1680
      Width           =   2055
   End
   Begin VB.TextBox txtSIDwrite 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Text            =   "00:00:00:00:00:00"
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label4 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      Caption         =   "MAC"
      Height          =   255
      Left            =   240
      TabIndex        =   80
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label32 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      Caption         =   "DAC-IFS"
      Height          =   255
      Left            =   5640
      TabIndex        =   13
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      Caption         =   "GID"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      Caption         =   "E-Key"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      Caption         =   "SID"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   735
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strHexDump As String
Dim bCapture As Boolean
Dim OwnMACAddress(7) As Byte
Dim cbuff(2000) As Byte
Dim SID_Node As Node
Dim fme As Integer

Const CURRENT = 1
Const DEFAULT = 2

Const UPSTREAM = 1
Const DOWNSTREAM = 2

Private Sub Form_Load()
Get_Parameter_Req
   
   txtMACAddress.Text = GetMACAddress
   
   OwnMACAddress(0) = "&H" & Mid(txtMACAddress.Text, 1, 2)
   OwnMACAddress(1) = "&H" & Mid(txtMACAddress.Text, 4, 2)
   OwnMACAddress(2) = "&H" & Mid(txtMACAddress.Text, 7, 2)
   OwnMACAddress(3) = "&H" & Mid(txtMACAddress.Text, 10, 2)
   OwnMACAddress(4) = "&H" & Mid(txtMACAddress.Text, 13, 2)
   OwnMACAddress(5) = "&H" & Mid(txtMACAddress.Text, 16, 2)

End Sub
Private Function Get_Parameter_Req()
    Dim tmpByteArray(59) As Byte
    Dim SelSID(7) As Byte
    'Get_Parameter_Req
    SelSID(0) = "&H" & Mid(Form1.ListSID.Text, 1, 2)
    SelSID(1) = "&H" & Mid(Form1.ListSID.Text, 4, 2)
    SelSID(2) = "&H" & Mid(Form1.ListSID.Text, 7, 2)
    SelSID(3) = "&H" & Mid(Form1.ListSID.Text, 10, 2)
    SelSID(4) = "&H" & Mid(Form1.ListSID.Text, 13, 2)
    SelSID(5) = "&H" & Mid(Form1.ListSID.Text, 16, 2)
    
    tmpByteArray(0) = SelSID(0): tmpByteArray(1) = SelSID(1): tmpByteArray(2) = SelSID(2)
    tmpByteArray(3) = SelSID(3): tmpByteArray(4) = SelSID(4): tmpByteArray(5) = SelSID(5)
    tmpByteArray(6) = OwnMACAddress(0) '&H0
    tmpByteArray(7) = OwnMACAddress(1) '&H23
    tmpByteArray(8) = OwnMACAddress(2) '&H26
    tmpByteArray(9) = OwnMACAddress(3) '&H37
    tmpByteArray(10) = OwnMACAddress(4) '&H85
    tmpByteArray(11) = OwnMACAddress(5) '&H6E
    tmpByteArray(12) = &H88: tmpByteArray(13) = &HC9
    tmpByteArray(14) = &H0: tmpByteArray(15) = &H1
    tmpByteArray(16) = &H0: tmpByteArray(17) = &H2 'SeqNum
    tmpByteArray(18) = &H0: tmpByteArray(19) = &H6
    tmpByteArray(20) = &H0: tmpByteArray(21) = &H0: tmpByteArray(22) = &H0: tmpByteArray(23) = &H0: tmpByteArray(24) = &H0
    tmpByteArray(25) = &H0: tmpByteArray(26) = &H0: tmpByteArray(27) = &H0: tmpByteArray(28) = &H0: tmpByteArray(29) = &H0
    tmpByteArray(30) = &H0: tmpByteArray(31) = &H0: tmpByteArray(32) = &H0: tmpByteArray(33) = &H0: tmpByteArray(34) = &H0
    tmpByteArray(35) = &H0: tmpByteArray(36) = &H0: tmpByteArray(37) = &H0: tmpByteArray(38) = &H0: tmpByteArray(39) = &H0
    tmpByteArray(40) = &H0: tmpByteArray(41) = &H0: tmpByteArray(42) = &H0: tmpByteArray(43) = &H0: tmpByteArray(44) = &H0
    tmpByteArray(45) = &H0: tmpByteArray(46) = &H0: tmpByteArray(47) = &H0: tmpByteArray(48) = &H0: tmpByteArray(49) = &H0
    tmpByteArray(50) = &H0: tmpByteArray(51) = &H0: tmpByteArray(52) = &H0: tmpByteArray(53) = &H0: tmpByteArray(54) = &H0
    tmpByteArray(55) = &H0: tmpByteArray(56) = &H0: tmpByteArray(57) = &H0: tmpByteArray(58) = &H0: tmpByteArray(59) = &H0
    
    vpSendPacket tmpByteArray()
End Function
Private Function GetMACAddress() As String
    Dim obj, objs
    
    Set objs = GetObject("winmgmts:").ExecQuery("SELECT MACAddress FROM Win32_NetworkAdapter WHERE MACAddress Is Not NULL")

    For Each obj In objs
        GetMACAddress = obj.MACAddress
        Exit For
    Next obj
End Function

Private Function GetId(Start As Integer, GetLen As Integer) As String
   Dim strID As String
   Dim i As Integer
    For i = Start To Start + GetLen - 1
       If cbuff(i) < &H10 Then
           strID = strID & ":0" & Hex(cbuff(i))
       Else
           strID = strID & ":" & Hex(cbuff(i))
       End If
    Next i
    strID = Right(strID, Len(strID) - 1)
    
    GetId = strID
End Function

Private Function Parse_Get_Parameter_Res()
    Dim dValue As Double
    
    '"Program Version"
    Form1.MSFlexGridParameter.TextMatrix(1, CURRENT) = "Ver" & Str((cbuff(49) * &H100 + cbuff(48)) / 100)
    Form1.MSFlexGridParameter.TextMatrix(1, DEFAULT) = "Ver" & Str((cbuff(21) * &H100 + cbuff(20)) / 100)
    
    '"Sub-code Version"
    Form1.MSFlexGridParameter.TextMatrix(2, CURRENT) = "Ver" & Val((cbuff(51) * &H100 + cbuff(50)) / 1000)
    Form1.MSFlexGridParameter.TextMatrix(2, DEFAULT) = "Ver" & Val((cbuff(51) * &H100 + cbuff(50)) / 1000) '??
    
    '"Station ID"
    Form1.MSFlexGridParameter.TextMatrix(3, CURRENT) = GetId(54, 6): Form1.MSFlexGridParameter.TextMatrix(3, DEFAULT) = GetId(31, 6)
    txtSIDwrite.Text = GetId(54, 6)
    '"Group ID"
    Form1.MSFlexGridParameter.TextMatrix(4, CURRENT) = GetId(110, 6): Form1.MSFlexGridParameter.TextMatrix(4, DEFAULT) = GetId(110, 6)
    txtGIDwrite.Text = GetId(110, 6)
    '"Device Type"
    Select Case (cbuff(53) And &H3)
        Case 0: Form1.MSFlexGridParameter.TextMatrix(5, CURRENT) = "Default"
        Case 1: Form1.MSFlexGridParameter.TextMatrix(5, CURRENT) = "Master"
        Case 2: Form1.MSFlexGridParameter.TextMatrix(5, CURRENT) = "Slave"
        Case 3: Form1.MSFlexGridParameter.TextMatrix(5, CURRENT) = "MeterGateway"
    End Select

    Select Case (cbuff(37) And &H3)
        Case 0: Form1.MSFlexGridParameter.TextMatrix(5, DEFAULT) = "Default"
        Case 1: Form1.MSFlexGridParameter.TextMatrix(5, DEFAULT) = "Master"
        Case 2: Form1.MSFlexGridParameter.TextMatrix(5, DEFAULT) = "Slave"
        Case 3: Form1.MSFlexGridParameter.TextMatrix(5, DEFAULT) = "MeterGateway"
    End Select
    
    '"Operation Mode"
    Select Case ((cbuff(75) And &HC) / &H100)
        Case 0: Form1.MSFlexGridParameter.TextMatrix(6, CURRENT) = "Active Mode"
        Case 1: Form1.MSFlexGridParameter.TextMatrix(6, CURRENT) = "Suspend Mode"
        Case 2: Form1.MSFlexGridParameter.TextMatrix(6, CURRENT) = "Factory Mode"
        Case 3: Form1.MSFlexGridParameter.TextMatrix(6, CURRENT) = "Programming Mode"
    End Select

    Select Case ((cbuff(39) And &HC) / &H100)
        Case 0: Form1.MSFlexGridParameter.TextMatrix(6, DEFAULT) = "Active Mode"
        Case 1: Form1.MSFlexGridParameter.TextMatrix(6, DEFAULT) = "Suspend Mode"
        Case 2: Form1.MSFlexGridParameter.TextMatrix(6, DEFAULT) = "Factory Mode"
        Case 3: Form1.MSFlexGridParameter.TextMatrix(6, DEFAULT) = "Programming Mode"
    End Select
    
    '"Serial Parameter" d:42 43 9600 8 1 None
    Dim strSerialSet As String
    
    dValue = cbuff(79) * &H100 + cbuff(78)
    strSerialSet = Val(1200 * 2 ^ ((dValue And &H1C0) / &H40)) 'Baudrate
    strSerialSet = strSerialSet & " " & Val(5 + ((dValue And &H30) / &H10)) ' Data Bit
    strSerialSet = strSerialSet & " " & Val(1 + ((dValue And &H8) / &H8)) ' Stop Bit
    If ((dValue And &H4) / &H4) = 0 Then
      strSerialSet = strSerialSet & " " & "None" 'Parity
    Else
        Select Case (dValue And &H3)
            Case 0: strSerialSet = strSerialSet & " " & "Odd"
            Case 1: strSerialSet = strSerialSet & " " & "Even"
            Case 2: strSerialSet = strSerialSet & " " & "1Set"
            Case 3: strSerialSet = strSerialSet & " " & "0Set"
        End Select
    End If
    Form1.MSFlexGridParameter.TextMatrix(7, CURRENT) = strSerialSet

    dValue = cbuff(43) * &H100 + cbuff(42)
    strSerialSet = Val(1200 * 2 ^ ((dValue And &H1C0) / &H40)) 'Baudrate
    strSerialSet = strSerialSet & " " & Val(5 + ((dValue And &H30) / &H10)) ' Data Bit
    strSerialSet = strSerialSet & " " & Val(1 + ((dValue And &H8) / &H8)) ' Stop Bit
    If ((dValue And &H4) / &H4) = 0 Then
      strSerialSet = strSerialSet & " " & "None" 'Parity
    Else
        Select Case (dValue And &H3)
            Case 0: strSerialSet = strSerialSet & " " & "Odd"
            Case 1: strSerialSet = strSerialSet & " " & "Even"
            Case 2: strSerialSet = strSerialSet & " " & "1Set"
            Case 3: strSerialSet = strSerialSet & " " & "0Set"
        End Select
    End If
    Form1.MSFlexGridParameter.TextMatrix(7, DEFAULT) = strSerialSet
    
    '"Self Reset Period"
    dValue = cbuff(75) * &H100 + cbuff(74)
    Form1.MSFlexGridParameter.TextMatrix(8, CURRENT) = Val((dValue And &H3C0) / &H40) & ":" & Val(dValue And &H3F) & "(Hour:Min)"
    dValue = cbuff(39) * &H100 + cbuff(38)
    Form1.MSFlexGridParameter.TextMatrix(8, DEFAULT) = Val((dValue And &H3C0) / &H40) & ":" & Val(dValue And &H3F) & "(Hour:Min)"
    
    '"2nd Interface Enable"
    If (cbuff(76) And &H20) = 0 Then
        Form1.MSFlexGridParameter.TextMatrix(9, CURRENT) = "Disable"
    Else
        Form1.MSFlexGridParameter.TextMatrix(9, CURRENT) = "Enable"
    End If
    
    If (cbuff(40) And &H20) = 0 Then
        Form1.MSFlexGridParameter.TextMatrix(9, DEFAULT) = "Disable"
    Else
        Form1.MSFlexGridParameter.TextMatrix(9, DEFAULT) = "Enable"
    End If
    
    ' "RTS CTS Status"
    If (cbuff(76) And &H10) = 0 Then
        Form1.MSFlexGridParameter.TextMatrix(10, CURRENT) = "Disable"
    Else
        Form1.MSFlexGridParameter.TextMatrix(10, CURRENT) = "Enable"
    End If
    
    If (cbuff(40) And &H10) = 0 Then
        Form1.MSFlexGridParameter.TextMatrix(10, DEFAULT) = "Disable"
    Else
        Form1.MSFlexGridParameter.TextMatrix(10, DEFAULT) = "Enable"
    End If
    
    '"Link Restriction"
    If (cbuff(75) And &H10) = 0 Then
        Form1.MSFlexGridParameter.TextMatrix(11, CURRENT) = "Disable"
    Else
        Form1.MSFlexGridParameter.TextMatrix(11, CURRENT) = "Enable"
    End If
    
    If (cbuff(39) And &H10) = 0 Then
        Form1.MSFlexGridParameter.TextMatrix(11, DEFAULT) = "Disable"
    Else
        Form1.MSFlexGridParameter.TextMatrix(11, DEFAULT) = "Enable"
    End If
    
    '"2nd Link Restriction"
    If (cbuff(76) And &H40) = 0 Then
        Form1.MSFlexGridParameter.TextMatrix(12, CURRENT) = "Disable"
    Else
        Form1.MSFlexGridParameter.TextMatrix(12, CURRENT) = "Enable"
    End If
    
    If (cbuff(40) And &H40) = 0 Then
        Form1.MSFlexGridParameter.TextMatrix(12, DEFAULT) = "Disable"
    Else
        Form1.MSFlexGridParameter.TextMatrix(12, DEFAULT) = "Enable"
    End If
    
    '"Tx Notification"
    If (cbuff(76) And &H8) = 0 Then
        Form1.MSFlexGridParameter.TextMatrix(13, CURRENT) = "Disable"
    Else
        Form1.MSFlexGridParameter.TextMatrix(13, CURRENT) = "Enable"
    End If
    
    If (cbuff(40) And &H8) = 0 Then
        Form1.MSFlexGridParameter.TextMatrix(13, DEFAULT) = "Disable"
    Else
        Form1.MSFlexGridParameter.TextMatrix(13, DEFAULT) = "Enable"
    End If
    
    '"Key Transmission"
    If (cbuff(76) And &H1) = 0 Then
        Form1.MSFlexGridParameter.TextMatrix(14, CURRENT) = "Disable"
    Else
        Form1.MSFlexGridParameter.TextMatrix(14, CURRENT) = "Enable"
    End If
    
    If (cbuff(40) And &H1) = 0 Then
        Form1.MSFlexGridParameter.TextMatrix(14, DEFAULT) = "Disable"
    Else
        Form1.MSFlexGridParameter.TextMatrix(14, DEFAULT) = "Enable"
    End If
    
    '"Cumulate MSDU number"
    Form1.MSFlexGridParameter.TextMatrix(15, CURRENT) = Val((cbuff(75) And &H60) / &H20)
    Form1.MSFlexGridParameter.TextMatrix(15, DEFAULT) = Val((cbuff(39) And &H60) / &H20)
    
    '"FBB TTL number"
    dValue = (((cbuff(77) And &H1F) * &H100 + cbuff(76)) And &H1F80) / &H80
    Form1.MSFlexGridParameter.TextMatrix(16, CURRENT) = Val(dValue)
    Form1.txtFBBTTLwrite = Val(dValue)
    dValue = (((cbuff(41) And &H1F) * &H100 + cbuff(40)) And &H1F80) / &H80
    Form1.MSFlexGridParameter.TextMatrix(16, DEFAULT) = Val(dValue)
    
    '"Firmware Upgrade Info."
    Form1.MSFlexGridParameter.TextMatrix(17, CURRENT) = "": Form1.MSFlexGridParameter.TextMatrix(17, DEFAULT) = ""
    
    '"Sub-Code Upgrade Info."
    Form1.MSFlexGridParameter.TextMatrix(18, CURRENT) = "": Form1.MSFlexGridParameter.TextMatrix(18, DEFAULT) = ""
    
    '"Tx FBB Status"
    Form1.MSFlexGridParameter.TextMatrix(19, CURRENT) = "": Form1.MSFlexGridParameter.TextMatrix(19, DEFAULT) = ""
    
    '"Tx Filter enable"
    Form1.MSFlexGridParameter.TextMatrix(20, CURRENT) = "": Form1.MSFlexGridParameter.TextMatrix(20, DEFAULT) = ""
    
    '"Parent Station ID"
    Form1.MSFlexGridParameter.TextMatrix(21, CURRENT) = GetId(98, 6): Form1.MSFlexGridParameter.TextMatrix(21, DEFAULT) = GetId(98, 6)
    Form1.txtParentSIDwrite.Text = GetId(98, 6)
    '"Parent BPS"
    Form1.MSFlexGridParameter.TextMatrix(22, CURRENT) = Val(cbuff(105) * &H100 + cbuff(104))
    Form1.MSFlexGridParameter.TextMatrix(22, DEFAULT) = Val(cbuff(105) * &H100 + cbuff(104)) '??
    Form1.txtParentBPS.Text = Val(cbuff(105) * &H100 + cbuff(104))
    '"Encryption Key"
    Form1.MSFlexGridParameter.TextMatrix(23, CURRENT) = GetId(60, 7): Form1.MSFlexGridParameter.TextMatrix(23, DEFAULT) = GetId(24, 7)
    txtKEYwrite.Text = GetId(60, 7)
    
End Function

