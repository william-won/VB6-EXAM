VERSION 5.00
Begin VB.Form frmDataSet 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "Set Data InputBox"
   ClientHeight    =   11355
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13965
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
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
   ScaleHeight     =   11355
   ScaleWidth      =   13965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.Frame fraOCTET10 
      Caption         =   "Octet-String(8-bit HEX) : 8-Byte Current Association LN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   3540
      TabIndex        =   117
      Top             =   9900
      Width           =   6855
      Begin VB.ComboBox cmbOctet10_2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5730
         Style           =   2  '드롭다운 목록
         TabIndex        =   122
         Top             =   390
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.ComboBox cmbOctet10_1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3810
         Style           =   2  '드롭다운 목록
         TabIndex        =   120
         Top             =   390
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.TextBox txtOctet10 
         Alignment       =   2  '가운데 맞춤
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1350
         MaxLength       =   8
         TabIndex        =   118
         Top             =   390
         Width           =   1485
      End
      Begin VB.Label lblTitle 
         Caption         =   "[Context]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   36
         Left            =   4800
         TabIndex        =   123
         Top             =   420
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label lblTitle 
         Caption         =   "[Assoc.]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   35
         Left            =   2940
         TabIndex        =   121
         Top             =   420
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblTitle 
         Caption         =   "[Password]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   34
         Left            =   240
         TabIndex        =   119
         Top             =   390
         Width           =   1215
      End
   End
   Begin VB.Frame fraOCTET16 
      Caption         =   "Octet-String(8-bit HEX) : 16-Byte Logical Device Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   3540
      TabIndex        =   114
      Top             =   8970
      Width           =   6855
      Begin VB.TextBox txtOctet16 
         Alignment       =   2  '가운데 맞춤
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1770
         MaxLength       =   16
         TabIndex        =   115
         Top             =   390
         Width           =   2415
      End
      Begin VB.Label lblTitle 
         Caption         =   "[Device Name]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   33
         Left            =   240
         TabIndex        =   116
         Top             =   390
         Width           =   1455
      End
   End
   Begin VB.Frame fraENUM 
      Caption         =   "ENUM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   6990
      TabIndex        =   41
      Top             =   10320
      Width           =   6855
      Begin VB.ComboBox cmbENUM 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   930
         Style           =   2  '드롭다운 목록
         TabIndex        =   94
         Top             =   360
         Width           =   885
      End
      Begin VB.Label lblTitle 
         Caption         =   "[HEX]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   23
         Left            =   240
         TabIndex        =   93
         Top             =   390
         Width           =   585
      End
   End
   Begin VB.Frame fraUI16 
      Caption         =   "Long-Unsigned(Unsigned16)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   120
      TabIndex        =   40
      Top             =   10320
      Width           =   6855
      Begin VB.TextBox txtUI16 
         Alignment       =   1  '오른쪽 맞춤
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1950
         MaxLength       =   5
         TabIndex        =   68
         Text            =   "0"
         Top             =   390
         Width           =   1485
      End
      Begin VB.Label lblTitle 
         Caption         =   "[Range] 0~65535"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   3720
         TabIndex        =   73
         Top             =   540
         Width           =   3015
      End
      Begin VB.Label lblUI16 
         Caption         =   "00 00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   4410
         TabIndex        =   71
         Top             =   270
         Width           =   615
      End
      Begin VB.Label lblTitle 
         Caption         =   "[HEX]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   3720
         TabIndex        =   70
         Top             =   240
         Width           =   585
      End
      Begin VB.Label lblTitle 
         Caption         =   "[2-Byte Unsigned]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   69
         Top             =   390
         Width           =   1725
      End
   End
   Begin VB.Frame fraUB08 
      Caption         =   "Unsigned(Unsigned8)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   6990
      TabIndex        =   39
      Top             =   9390
      Width           =   6855
      Begin VB.TextBox txtUB08 
         Alignment       =   1  '오른쪽 맞춤
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1950
         MaxLength       =   3
         TabIndex        =   88
         Text            =   "0"
         Top             =   390
         Width           =   1485
      End
      Begin VB.Label lblTitle 
         Caption         =   "[Range] 0~255"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   22
         Left            =   3720
         TabIndex        =   92
         Top             =   540
         Width           =   3015
      End
      Begin VB.Label lblUB08 
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   4410
         TabIndex        =   91
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblTitle 
         Caption         =   "[HEX]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   21
         Left            =   3720
         TabIndex        =   90
         Top             =   240
         Width           =   585
      End
      Begin VB.Label lblTitle 
         Caption         =   "[1-Byte Unsigned]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   20
         Left            =   240
         TabIndex        =   89
         Top             =   390
         Width           =   1725
      End
   End
   Begin VB.Frame fraSI16 
      Caption         =   "Long(Integer16)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   120
      TabIndex        =   38
      Top             =   9390
      Width           =   6855
      Begin VB.TextBox txtSI16 
         Alignment       =   1  '오른쪽 맞춤
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1950
         MaxLength       =   6
         TabIndex        =   64
         Text            =   "0"
         Top             =   390
         Width           =   1485
      End
      Begin VB.Label lblTitle 
         Caption         =   "[Range] -32768~32767"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   3720
         TabIndex        =   72
         Top             =   540
         Width           =   3015
      End
      Begin VB.Label lblSI16 
         Caption         =   "00 00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   4410
         TabIndex        =   67
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblTitle 
         Caption         =   "[HEX]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   3720
         TabIndex        =   66
         Top             =   240
         Width           =   585
      End
      Begin VB.Label lblTitle 
         Caption         =   "[2-Byte Signed]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   65
         Top             =   390
         Width           =   1725
      End
   End
   Begin VB.Frame fraSB08 
      Caption         =   "Integer(Integer8)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   6990
      TabIndex        =   37
      Top             =   8460
      Width           =   6855
      Begin VB.TextBox txtSB08 
         Alignment       =   1  '오른쪽 맞춤
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1950
         MaxLength       =   4
         TabIndex        =   83
         Text            =   "0"
         Top             =   390
         Width           =   1485
      End
      Begin VB.Label lblTitle 
         Caption         =   "[Range] -128~127"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   19
         Left            =   3720
         TabIndex        =   87
         Top             =   540
         Width           =   3015
      End
      Begin VB.Label lblSB08 
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   4410
         TabIndex        =   86
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblTitle 
         Caption         =   "[HEX]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   18
         Left            =   3720
         TabIndex        =   85
         Top             =   240
         Width           =   585
      End
      Begin VB.Label lblTitle 
         Caption         =   "[1-Byte Signed]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   17
         Left            =   240
         TabIndex        =   84
         Top             =   390
         Width           =   1725
      End
   End
   Begin VB.Frame fraOCTET01 
      Caption         =   "Octet-String(8-bit HEX) : 1-Byte Hex"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   120
      TabIndex        =   36
      Top             =   8460
      Width           =   6855
      Begin VB.ComboBox cmbOctet01 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   930
         Style           =   2  '드롭다운 목록
         TabIndex        =   63
         Top             =   360
         Width           =   885
      End
      Begin VB.Label lblTitle 
         Caption         =   "[HEX]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   62
         Top             =   390
         Width           =   585
      End
   End
   Begin VB.Frame fraOCTET04 
      Caption         =   "Octet-String(8-bit HEX) : 4-Byte Float"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   6990
      TabIndex        =   35
      Top             =   7530
      Width           =   6855
      Begin VB.TextBox txtOctet04 
         Alignment       =   1  '오른쪽 맞춤
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1950
         MaxLength       =   10
         TabIndex        =   79
         Text            =   "1"
         Top             =   390
         Width           =   1485
      End
      Begin VB.Label lblTitle 
         Caption         =   "[4-Byte Float]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   16
         Left            =   240
         TabIndex        =   82
         Top             =   390
         Width           =   1245
      End
      Begin VB.Label lblOctet04 
         Caption         =   "3F 80 00 00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   3720
         TabIndex        =   81
         Top             =   540
         Width           =   1245
      End
      Begin VB.Label lblTitle 
         Caption         =   "[HEX, 4321 order]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   15
         Left            =   3720
         TabIndex        =   80
         Top             =   240
         Width           =   1665
      End
   End
   Begin VB.Frame fraOCTET07 
      Caption         =   "Octet-String(8-bit HEX) : 7-Byte Meter ID"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   120
      TabIndex        =   34
      Top             =   7530
      Width           =   6855
      Begin VB.TextBox txtOctet07 
         Alignment       =   2  '가운데 맞춤
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1770
         MaxLength       =   7
         TabIndex        =   59
         Text            =   "0000000"
         Top             =   390
         Width           =   1665
      End
      Begin VB.Label lblOctet07 
         Caption         =   "30 30 30 30 30 30 30"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   4410
         TabIndex        =   61
         Top             =   420
         Width           =   2205
      End
      Begin VB.Label lblTitle 
         Caption         =   "[HEX]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   3720
         TabIndex        =   60
         Top             =   390
         Width           =   585
      End
      Begin VB.Label lblTitle 
         Caption         =   "[Meter ID]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   58
         Top             =   390
         Width           =   1245
      End
   End
   Begin VB.Frame fraOCTET12 
      Caption         =   "Octet-String(8-bit HEX) : 12-Byte DateTime"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   6990
      TabIndex        =   33
      Top             =   6600
      Width           =   6855
      Begin VB.Timer tmrOctet12 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   6390
         Top             =   150
      End
      Begin VB.TextBox txtOctet12Time 
         Alignment       =   2  '가운데 맞춤
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   2
         Left            =   3270
         MaxLength       =   2
         TabIndex        =   106
         Text            =   "0"
         Top             =   480
         Width           =   435
      End
      Begin VB.TextBox txtOctet12Time 
         Alignment       =   2  '가운데 맞춤
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   1
         Left            =   2850
         MaxLength       =   2
         TabIndex        =   104
         Text            =   "0"
         Top             =   480
         Width           =   435
      End
      Begin VB.TextBox txtOctet12Time 
         Alignment       =   2  '가운데 맞춤
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   0
         Left            =   2430
         MaxLength       =   2
         TabIndex        =   102
         Text            =   "0"
         Top             =   480
         Width           =   435
      End
      Begin VB.TextBox txtOctet12Date 
         Alignment       =   2  '가운데 맞춤
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   2
         Left            =   1950
         MaxLength       =   2
         TabIndex        =   100
         Text            =   "1"
         Top             =   480
         Width           =   435
      End
      Begin VB.TextBox txtOctet12Date 
         Alignment       =   2  '가운데 맞춤
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   1
         Left            =   1530
         MaxLength       =   2
         TabIndex        =   98
         Text            =   "1"
         Top             =   480
         Width           =   435
      End
      Begin VB.TextBox txtOctet12Date 
         Alignment       =   2  '가운데 맞춤
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   0
         Left            =   960
         MaxLength       =   4
         TabIndex        =   96
         Text            =   "2005"
         Top             =   480
         Width           =   585
      End
      Begin VB.CheckBox chkOctet12 
         Caption         =   "PC Time"
         Height          =   495
         Left            =   150
         TabIndex        =   95
         Top             =   300
         Value           =   1  '확인
         Width           =   825
      End
      Begin VB.Label lblOctet12Time 
         Caption         =   "00 00 00 00 00 00 00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   4830
         TabIndex        =   111
         Top             =   570
         Width           =   1935
      End
      Begin VB.Label lblTitle 
         Caption         =   "[HMSmdvT]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   31
         Left            =   3780
         TabIndex        =   110
         Top             =   540
         Width           =   1095
      End
      Begin VB.Label lblTitle 
         Caption         =   "[YYMDW]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   30
         Left            =   3780
         TabIndex        =   109
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label lblOctet12Date 
         Caption         =   "00 00 00 00 00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   4830
         TabIndex        =   108
         Top             =   270
         Width           =   1455
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  '가운데 맞춤
         Caption         =   "sec."
         Height          =   255
         Index           =   29
         Left            =   3270
         TabIndex        =   107
         Top             =   240
         Width           =   465
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  '가운데 맞춤
         Caption         =   "min."
         Height          =   255
         Index           =   28
         Left            =   2850
         TabIndex        =   105
         Top             =   240
         Width           =   465
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  '가운데 맞춤
         Caption         =   "hour"
         Height          =   255
         Index           =   27
         Left            =   2430
         TabIndex        =   103
         Top             =   240
         Width           =   465
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  '가운데 맞춤
         Caption         =   "day"
         Height          =   255
         Index           =   26
         Left            =   1950
         TabIndex        =   101
         Top             =   240
         Width           =   465
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  '가운데 맞춤
         Caption         =   "mon."
         Height          =   255
         Index           =   25
         Left            =   1530
         TabIndex        =   99
         Top             =   240
         Width           =   465
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  '가운데 맞춤
         Caption         =   "year"
         Height          =   255
         Index           =   24
         Left            =   960
         TabIndex        =   97
         Top             =   240
         Width           =   585
      End
   End
   Begin VB.Frame fraOCTET08 
      Caption         =   "Octet-String(8-bit HEX) : 8-Byte Program ID"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   120
      TabIndex        =   32
      Top             =   6600
      Width           =   6855
      Begin VB.TextBox txtOctet08 
         Alignment       =   2  '가운데 맞춤
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1770
         MaxLength       =   8
         TabIndex        =   54
         Text            =   "330v2A  "
         Top             =   390
         Width           =   1665
      End
      Begin VB.Label lblTitle 
         Caption         =   "[Program ID]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   57
         Top             =   390
         Width           =   1245
      End
      Begin VB.Label lblOctet08 
         Caption         =   "33 33 30 76 32 41 20 20"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   4410
         TabIndex        =   56
         Top             =   420
         Width           =   2205
      End
      Begin VB.Label lblTitle 
         Caption         =   "[HEX]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   3720
         TabIndex        =   55
         Top             =   390
         Width           =   585
      End
   End
   Begin VB.Frame fraUL32 
      Caption         =   "Double-Long-Unsigned(Unsigned32)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   6990
      TabIndex        =   31
      Top             =   5670
      Width           =   6855
      Begin VB.TextBox txtUL32 
         Alignment       =   1  '오른쪽 맞춤
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1950
         MaxLength       =   10
         TabIndex        =   74
         Text            =   "0"
         Top             =   390
         Width           =   1485
      End
      Begin VB.Label lblUL32 
         Caption         =   "00 00 00 00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   4410
         TabIndex        =   78
         Top             =   270
         Width           =   1155
      End
      Begin VB.Label lblTitle 
         Caption         =   "[Range] 0~4294967295"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   14
         Left            =   3720
         TabIndex        =   77
         Top             =   540
         Width           =   3015
      End
      Begin VB.Label lblTitle 
         Caption         =   "[HEX]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   3720
         TabIndex        =   76
         Top             =   240
         Width           =   585
      End
      Begin VB.Label lblTitle 
         Caption         =   "[4-Byte Unsigned]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   240
         TabIndex        =   75
         Top             =   390
         Width           =   1725
      End
   End
   Begin VB.Frame fraBITSTR 
      Caption         =   "Bit-String"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   120
      TabIndex        =   30
      Top             =   5670
      Width           =   6855
      Begin VB.CheckBox chkBITSTR 
         Caption         =   "bit0"
         Height          =   345
         Index           =   7
         Left            =   5280
         TabIndex        =   51
         Top             =   360
         Value           =   1  '확인
         Width           =   645
      End
      Begin VB.CheckBox chkBITSTR 
         Caption         =   "bit1"
         Height          =   345
         Index           =   6
         Left            =   4560
         TabIndex        =   50
         Top             =   360
         Width           =   645
      End
      Begin VB.CheckBox chkBITSTR 
         Caption         =   "bit2"
         Height          =   345
         Index           =   5
         Left            =   3840
         TabIndex        =   49
         Top             =   360
         Width           =   645
      End
      Begin VB.CheckBox chkBITSTR 
         Caption         =   "bit3"
         Height          =   345
         Index           =   4
         Left            =   3120
         TabIndex        =   48
         Top             =   360
         Width           =   645
      End
      Begin VB.CheckBox chkBITSTR 
         Caption         =   "bit4"
         Height          =   345
         Index           =   3
         Left            =   2400
         TabIndex        =   47
         Top             =   360
         Width           =   645
      End
      Begin VB.CheckBox chkBITSTR 
         Caption         =   "bit5"
         Height          =   345
         Index           =   2
         Left            =   1680
         TabIndex        =   46
         Top             =   360
         Width           =   645
      End
      Begin VB.CheckBox chkBITSTR 
         Caption         =   "bit6"
         Height          =   345
         Index           =   1
         Left            =   960
         TabIndex        =   45
         Top             =   360
         Width           =   645
      End
      Begin VB.CheckBox chkBITSTR 
         Caption         =   "bit7"
         Height          =   345
         Index           =   0
         Left            =   270
         TabIndex        =   44
         Top             =   360
         Width           =   645
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  '가운데 맞춤
         Caption         =   "[HEX]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   6030
         TabIndex        =   53
         Top             =   270
         Width           =   645
      End
      Begin VB.Label lblBitStr 
         Alignment       =   2  '가운데 맞춤
         Caption         =   "01"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   6150
         TabIndex        =   52
         Top             =   540
         Width           =   405
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   9450
      TabIndex        =   29
      Top             =   4950
      Width           =   1965
   End
   Begin VB.CommandButton cmdSet 
      Caption         =   "Accept &Set"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   7230
      TabIndex        =   28
      Top             =   4950
      Width           =   1965
   End
   Begin VB.Frame fraBool 
      Caption         =   "Boolean"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   120
      TabIndex        =   27
      Top             =   4740
      Width           =   6855
      Begin VB.OptionButton optBOOL 
         Caption         =   "TRUE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   1650
         TabIndex        =   43
         Top             =   360
         Width           =   1155
      End
      Begin VB.OptionButton optBOOL 
         Caption         =   "FALSE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   240
         TabIndex        =   42
         Top             =   360
         Value           =   -1  'True
         Width           =   1155
      End
      Begin VB.Label lblBOOL 
         Alignment       =   2  '가운데 맞춤
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   4440
         TabIndex        =   113
         Top             =   390
         Width           =   405
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  '가운데 맞춤
         Caption         =   "[HEX]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   32
         Left            =   3720
         TabIndex        =   112
         Top             =   390
         Width           =   645
      End
   End
   Begin VB.Frame fraDataType 
      Caption         =   "Information for Data Type Selected"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4425
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   11385
      Begin VB.OptionButton optType 
         Caption         =   "255 : Don't Care(Null), N-Byte"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   315
         Index           =   25
         Left            =   180
         TabIndex        =   26
         Top             =   3960
         Width           =   5205
      End
      Begin VB.OptionButton optType 
         Caption         =   "24 : Float64(Octet-String,Size(8)), 8-Byte"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   315
         Index           =   24
         Left            =   5880
         TabIndex        =   25
         Top             =   3600
         Width           =   5205
      End
      Begin VB.OptionButton optType 
         Caption         =   "23 : Float32(Octet-String,Size(4)), 4-Byte"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   315
         Index           =   23
         Left            =   180
         TabIndex        =   24
         Top             =   3600
         Width           =   5205
      End
      Begin VB.OptionButton optType 
         Caption         =   "22 : ENUM, 1-Byte"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   22
         Left            =   5880
         TabIndex        =   23
         Top             =   3240
         Width           =   5205
      End
      Begin VB.OptionButton optType 
         Caption         =   "21 : Long64-Unsigned(Unsigned64), 8-Byte"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   315
         Index           =   21
         Left            =   180
         TabIndex        =   22
         Top             =   3240
         Width           =   5205
      End
      Begin VB.OptionButton optType 
         Caption         =   "20 : Long64(Integer64), 8-Byte"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   315
         Index           =   20
         Left            =   5880
         TabIndex        =   21
         Top             =   2880
         Width           =   5205
      End
      Begin VB.OptionButton optType 
         Caption         =   "19 : Compact-Array(Sequence), N-Byte"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   315
         Index           =   19
         Left            =   180
         TabIndex        =   20
         Top             =   2880
         Width           =   5205
      End
      Begin VB.OptionButton optType 
         Caption         =   "18 : Long-Unsigned(Unsigned16), 2-Byte"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   18
         Left            =   5880
         TabIndex        =   19
         Top             =   2520
         Width           =   5205
      End
      Begin VB.OptionButton optType 
         Caption         =   "17 : Unsigned(Unsigned8), 1-Byte"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   17
         Left            =   180
         TabIndex        =   18
         Top             =   2520
         Width           =   5205
      End
      Begin VB.OptionButton optType 
         Caption         =   "16 : Long(Integer16), 2-Byte"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   16
         Left            =   7620
         TabIndex        =   17
         Top             =   2160
         Width           =   3615
      End
      Begin VB.OptionButton optType 
         Caption         =   "15 : Integer(Integer8), 1-Byte"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   15
         Left            =   3900
         TabIndex        =   16
         Top             =   2160
         Width           =   3615
      End
      Begin VB.OptionButton optType 
         Caption         =   "14 : Not Defined"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   315
         Index           =   14
         Left            =   180
         TabIndex        =   15
         Top             =   2160
         Width           =   3615
      End
      Begin VB.OptionButton optType 
         Caption         =   "13 : Bcd(Integer8), 1-Byte"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   315
         Index           =   13
         Left            =   7620
         TabIndex        =   14
         Top             =   1800
         Width           =   3615
      End
      Begin VB.OptionButton optType 
         Caption         =   "12 : Not Defined"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   315
         Index           =   12
         Left            =   3900
         TabIndex        =   13
         Top             =   1800
         Width           =   3615
      End
      Begin VB.OptionButton optType 
         Caption         =   "11 : Time(GeneralizedTime), N-Byte"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   315
         Index           =   11
         Left            =   180
         TabIndex        =   12
         Top             =   1800
         Width           =   3615
      End
      Begin VB.OptionButton optType 
         Caption         =   "10 : Visible-String(ASCII), N-Byte"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   315
         Index           =   10
         Left            =   7620
         TabIndex        =   11
         Top             =   1440
         Width           =   3615
      End
      Begin VB.OptionButton optType 
         Caption         =   "09 : Octet-String(8-bit HEX), N-Byte"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   315
         Index           =   9
         Left            =   3900
         TabIndex        =   10
         Top             =   1440
         Width           =   3615
      End
      Begin VB.OptionButton optType 
         Caption         =   "08 : Not Defined"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   315
         Index           =   8
         Left            =   180
         TabIndex        =   9
         Top             =   1440
         Width           =   3615
      End
      Begin VB.OptionButton optType 
         Caption         =   "05 : Double-Long(Integer32), 4-Byte"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   315
         Index           =   5
         Left            =   7620
         TabIndex        =   8
         Top             =   720
         Width           =   3615
      End
      Begin VB.OptionButton optType 
         Caption         =   "04 : Bit-String, N-Byte"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   315
         Index           =   4
         Left            =   3900
         TabIndex        =   7
         Top             =   720
         Width           =   3615
      End
      Begin VB.OptionButton optType 
         Caption         =   "06 : Double-Long-Unsigned(Unsigned32), 4-Byte"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   6
         Left            =   180
         TabIndex        =   6
         Top             =   1080
         Width           =   5205
      End
      Begin VB.OptionButton optType 
         Caption         =   "02 : Structure(Sequence), N-Byte"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   2
         Left            =   7620
         TabIndex        =   5
         Top             =   360
         Width           =   3615
      End
      Begin VB.OptionButton optType 
         Caption         =   "01 : Array(Sequence), N-Byte"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   1
         Left            =   3900
         TabIndex        =   4
         Top             =   360
         Width           =   3615
      End
      Begin VB.OptionButton optType 
         Caption         =   "03 : Boolean, 1-Byte"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   3
         Left            =   180
         TabIndex        =   3
         Top             =   720
         Width           =   3615
      End
      Begin VB.OptionButton optType 
         Caption         =   "07 : Floating-Point(Octet-String,Size(4)), 4-Byte"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   315
         Index           =   7
         Left            =   5880
         TabIndex        =   2
         Top             =   1080
         Width           =   5205
      End
      Begin VB.OptionButton optType 
         Caption         =   "00 : Null Data(Null), 0-Byte"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   315
         Index           =   0
         Left            =   180
         TabIndex        =   1
         Top             =   360
         Width           =   3615
      End
   End
End
Attribute VB_Name = "frmDataSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*                                                                                                  *

Private Sub Form_Load()

    Call Init_Frame_Show
    
    Call DataType_Frame_Show
    
End Sub
'*                                                                                                  *

Private Sub Init_Frame_Show()

    Dim ii_Int As Integer
    
    Me.Height = 6270:   Me.Width = 11685

    Call Arrange_Form(fraBool, False):      Call Arrange_Form(fraBITSTR, False)
    Call Arrange_Form(fraUL32, False):      Call Arrange_Form(fraOCTET08, False)
    Call Arrange_Form(fraOCTET12, False):   Call Arrange_Form(fraOCTET07, False)
    Call Arrange_Form(fraOCTET04, False):   Call Arrange_Form(fraOCTET01, False)
    Call Arrange_Form(fraOCTET16, False):   Call Arrange_Form(fraOCTET10, False)
    Call Arrange_Form(fraSB08, False):      Call Arrange_Form(fraSI16, False)
    Call Arrange_Form(fraUB08, False):      Call Arrange_Form(fraUI16, False)
    Call Arrange_Form(fraENUM, False)
    
    With cmbOctet01
        .Clear
        For ii_Int = 0 To 255 Step 1
            .AddItem HexToTwo(Hex(ii_Int))
        Next ii_Int
        .ListIndex = 0
    End With
    With cmbENUM
        .Clear
        For ii_Int = 0 To 255 Step 1
            .AddItem HexToTwo(Hex(ii_Int))
        Next ii_Int
        .ListIndex = 1
    End With
    With cmbOctet10_1
        .Clear
        .AddItem 1:     .AddItem 2
        .AddItem 4:     .AddItem 5
        .ListIndex = 0
    End With
    With cmbOctet10_2
        .Clear
        .AddItem 1
        .AddItem 2
        .ListIndex = 0
    End With
    
End Sub
'*                                                                                                  *

Private Sub DataType_Frame_Show()

    optType(gSet_DataType).Value = 1
    optType(gSet_DataType).FontBold = True
    
    Select Case gSet_DataType
        Case 3:     fraBool.Visible = True      'Boolean
        Case 4:     fraBITSTR.Visible = True    'BitString
        Case 6:     fraUL32.Visible = True      'Unsigned32
        Case 15:    fraSB08.Visible = True      'Integer8
        Case 16:    fraSI16.Visible = True      'Integer16
        Case 17:    fraUB08.Visible = True      'Unsigned8
        Case 18:    fraUI16.Visible = True      'Unsigned16
        Case 22:    fraENUM.Visible = True      'ENUM
        Case 9
            Select Case gSet_DataLen
                Case 12
                    fraOCTET12.Visible = True           'Octet12(DateTime)
                    Call tmrOctet12_Timer
                    tmrOctet12.Enabled = True
                Case 8:     fraOCTET08.Visible = True   'Octet08(Program ID)
                Case 7:     fraOCTET07.Visible = True   'Octet07(Meter ID)
                Case 4:     fraOCTET04.Visible = True   'Octet04(Float)
                Case 1:     fraOCTET01.Visible = True   'Octet01(Hex)
                Case 16:    fraOCTET16.Visible = True   'Octet16(Logical Device Name)
                Case 10:    fraOCTET10.Visible = True   'Octet10(Password)
            End Select
        Case Else
            MsgBox "The datatype writing is not supported now...", vbInformation, "Not Supported"
            cmdSet.Enabled = False
    End Select

End Sub
'*                                                                                                  *

Private Sub Arrange_Form(ByRef myForm As Frame, ByVal bSHOW As Boolean)
    With myForm
        .Left = 120:    .Top = 4740
        .Visible = bSHOW
    End With
End Sub
'*                                                                                                  *

Private Sub cmdCancel_Click()
    Set_Confirmed_Or_Not = False
    Unload Me
End Sub
'*                                                                                                  *

Private Sub cmdSet_Click()

    Dim tARRAY() As String
    Dim ii_Byte As Byte
    Dim tSTR As String
        
    Set_Confirmed_Or_Not = False
    
    Select Case gSet_DataType
        Case 3      'Boolean
            ls_buf(0) = CByte("&H" & lblBOOL.Caption)
        Case 4      'BitString
            ls_buf(0) = CByte("&H" & lblBitStr.Caption)
        Case 6      'Unsigned32
            tARRAY = Split(lblUL32.Caption, " ")
            If UBound(tARRAY) = (gSet_DataLen - 1) Then
                For ii_Byte = 0 To 3 Step 1
                    ls_buf(ii_Byte) = CByte("&H" & tARRAY(ii_Byte))
                Next ii_Byte
            Else
                MsgBox "Datatype and datalength mismatch. Cannot write the value!", _
                        vbExclamation, "Wrong Data"
                Exit Sub
            End If
        Case 15     'Integer8
            ls_buf(0) = CByte("&H" & lblSB08.Caption)
        Case 16     'Integer16
            tARRAY = Split(lblSI16.Caption, " ")
            If UBound(tARRAY) = (gSet_DataLen - 1) Then
                For ii_Byte = 0 To 1 Step 1
                    ls_buf(ii_Byte) = CByte("&H" & tARRAY(ii_Byte))
                Next ii_Byte
            Else
                MsgBox "Datatype and datalength mismatch. Cannot write the value!", _
                        vbExclamation, "Wrong Data"
                Exit Sub
            End If
        Case 17     'Unsigned8
            ls_buf(0) = CByte("&H" & lblUB08.Caption)
        Case 18:    fraUI16.Visible = True      'Unsigned16
            tARRAY = Split(lblUI16.Caption, " ")
            If UBound(tARRAY) = (gSet_DataLen - 1) Then
                For ii_Byte = 0 To 1 Step 1
                    ls_buf(ii_Byte) = CByte("&H" & tARRAY(ii_Byte))
                Next ii_Byte
            Else
                MsgBox "Datatype and datalength mismatch. Cannot write the value!", _
                        vbExclamation, "Wrong Data"
                Exit Sub
            End If
        Case 22     'ENUM
            ls_buf(0) = cmbENUM.ListIndex
        Case 9
            Select Case gSet_DataLen
                Case 12     'Octet12(DateTime)
                    tARRAY = Split(lblOctet12Date.Caption & " " & lblOctet12Time.Caption, " ")
                    If UBound(tARRAY) = (gSet_DataLen - 1) Then
                        For ii_Byte = 0 To 11 Step 1
                            ls_buf(ii_Byte) = CByte("&H" & tARRAY(ii_Byte))
                        Next ii_Byte
                    Else
                        MsgBox "Datatype and datalength mismatch. Cannot write the value!", _
                                vbExclamation, "Wrong Data"
                        Exit Sub
                    End If
                Case 8      'Octet08(Program ID)
                    tARRAY = Split(lblOctet08.Caption, " ")
                    If UBound(tARRAY) = (gSet_DataLen - 1) Then
                        For ii_Byte = 0 To 7 Step 1
                            ls_buf(ii_Byte) = CByte("&H" & tARRAY(ii_Byte))
                        Next ii_Byte
                    Else
                        MsgBox "Datatype and datalength mismatch. Cannot write the value!", _
                                vbExclamation, "Wrong Data"
                        Exit Sub
                    End If
                Case 7      'Octet07(Meter ID)
                    tARRAY = Split(lblOctet07.Caption, " ")
                    If UBound(tARRAY) = (gSet_DataLen - 1) Then
                        For ii_Byte = 0 To 6 Step 1
                            ls_buf(ii_Byte) = CByte("&H" & tARRAY(ii_Byte))
                        Next ii_Byte
                    Else
                        MsgBox "Datatype and datalength mismatch. Cannot write the value!", _
                                vbExclamation, "Wrong Data"
                        Exit Sub
                    End If
                Case 4      'Octet04(Float)
                    tARRAY = Split(lblOctet04.Caption, " ")
                    If UBound(tARRAY) = (gSet_DataLen - 1) Then
                        For ii_Byte = 0 To 3 Step 1
                            ls_buf(ii_Byte) = CByte("&H" & tARRAY(ii_Byte))
                        Next ii_Byte
                    Else
                        MsgBox "Datatype and datalength mismatch. Cannot write the value!", _
                                vbExclamation, "Wrong Data"
                        Exit Sub
                    End If
                Case 1      'Octet01(Hex)
                    ls_buf(0) = cmbOctet01.ListIndex
                Case 16     'Octet16(Logical Device Name)
                    tSTR = txtOctet16.Text
                    If Len(tSTR) > 16 Then tSTR = Mid(tSTR, 1, 16)
                    For ii_Byte = 0 To 15 Step 1
                        If ii_Byte < Len(tSTR) Then
                            ls_buf(ii_Byte) = Asc(Mid(tSTR, ii_Byte + 1, 1))
                        Else
                            ls_buf(ii_Byte) = &H0
                        End If
                    Next ii_Byte
                Case 10     'Octet10(Current Association LN-Password)
'                    ls_buf(0) = &H10 + CByte(cmbOctet10_1.Text)
'                    ls_buf(1) = CByte(cmbOctet10_2.Text)
'                    tSTR = txtOctet10.Text
'                    If Len(tSTR) > 8 Then tSTR = Mid(tSTR, 1, 8)
'                    For ii_Byte = 0 To 7 Step 1
'                        If ii_Byte < Len(tSTR) Then
'                            ls_buf(2 + ii_Byte) = Asc(Mid(tSTR, ii_Byte + 1, 1))
'                        Else
'                            ls_buf(2 + ii_Byte) = &H0
'                        End If
'                    Next ii_Byte
                    tSTR = txtOctet10.Text
                    If Len(tSTR) > 8 Then tSTR = Mid(tSTR, 1, 8)
                    For ii_Byte = 0 To 7 Step 1
                        If ii_Byte < Len(tSTR) Then
                            ls_buf(ii_Byte) = Asc(Mid(tSTR, ii_Byte + 1, 1))
                        Else
                            ls_buf(ii_Byte) = &H20
                        End If
                    Next ii_Byte
                    gSet_DataLen = 8    'ID 2byte 제외
            End Select
        Case Else
            MsgBox "This datatype is not supported for writing the value!", _
                    vbExclamation, "Wrong DataType"
            Exit Sub
    End Select
    
    Erase tARRAY
    Set_Confirmed_Or_Not = True
        
    Unload Me
    
End Sub
'*                                                                                                  *


Private Sub optBOOL_Click(Index As Integer)
    If optBOOL(0).Value = True Then
        lblBOOL.Caption = "00"
    Else
        lblBOOL.Caption = "01"
    End If
End Sub
'*                                                                                                  *

Private Sub chkBITSTR_Click(Index As Integer)
    
    Dim tBYTE As Byte
    Dim ii_Byte As Byte
    
    For ii_Byte = 0 To 7 Step 1
        tBYTE = tBYTE + (IIf(chkBITSTR(ii_Byte).Value, 1, 0) * 2 ^ (7 - ii_Byte))
    Next ii_Byte
    
    lblBitStr.Caption = HexToTwo(Hex(tBYTE))
    
End Sub
'*                                                                                                  *

Private Sub txtOctet07_GotFocus()
    Call Select_TextBox_Str(txtOctet07)
End Sub
'*                                                                                                  *

Private Sub txtOctet07_LostFocus()
    
    Dim tSTR As String
    Dim tLEN As Byte
    
    With txtOctet07
        tLEN = Len(.Text)
        If tLEN < 7 Then
            .Text = String(7 - tLEN, "0") & .Text
        End If
    End With
    Call txtOctet07_KeyUp(0, 0)
    
End Sub
'*                                                                                                  *

Private Sub txtOctet07_KeyPress(KeyAscii As Integer)
    KeyAscii = Check_AsciiOnly(KeyAscii)
    With txtOctet07
        If Len(.Text) > 7 Then
            .Text = Mid(.Text, 1, 7)
        End If
    End With
End Sub
'*                                                                                                  *

Private Sub txtOctet07_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim ii_Byte As Byte
    Dim tSTR As String
    Dim tLEN As Byte
    
    With txtOctet07
        tLEN = Len(.Text)
        If tLEN = 0 Then
            lblOctet07.Caption = "00 00 00 00 00 00 00"
        Else
            For ii_Byte = 1 To tLEN Step 1
                tSTR = tSTR & HexToTwo(Hex(Asc(Mid(.Text, ii_Byte, 1)))) & " "
            Next ii_Byte
            If tLEN < 7 Then
                For ii_Byte = (tLEN + 1) To 7 Step 1
                    tSTR = tSTR & "00" & " "
                Next ii_Byte
            End If
            tSTR = Trim(tSTR)
        End If
    End With
    lblOctet07.Caption = tSTR

End Sub
'*                                                                                                  *

Private Sub txtOctet08_GotFocus()
    Call Select_TextBox_Str(txtOctet08)
End Sub
'*                                                                                                  *

Private Sub txtOctet08_LostFocus()
    
    Dim tSTR As String
    Dim tLEN As Byte
    
    With txtOctet08
        tLEN = Len(.Text)
        If tLEN < 8 Then
            .Text = String(8 - tLEN, " ") & .Text
        End If
    End With
    Call txtOctet08_KeyUp(0, 0)

End Sub
'*                                                                                                  *

Private Sub txtOctet08_KeyPress(KeyAscii As Integer)
    KeyAscii = Check_AsciiOnly(KeyAscii)
    With txtOctet08
        If Len(.Text) > 8 Then
            .Text = Mid(.Text, 1, 8)
        End If
    End With
End Sub
'*                                                                                                  *

Private Sub txtOctet08_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim ii_Byte As Byte
    Dim tSTR As String
    Dim tLEN As Byte
    
    With txtOctet08
        tLEN = Len(.Text)
        If tLEN = 0 Then
            lblOctet08.Caption = "00 00 00 00 00 00 00 00"
        Else
            For ii_Byte = 1 To tLEN Step 1
                tSTR = tSTR & HexToTwo(Hex(Asc(Mid(.Text, ii_Byte, 1)))) & " "
            Next ii_Byte
            If tLEN < 8 Then
                For ii_Byte = tLEN To 8 Step 1
                    tSTR = tSTR & "00" & " "
                Next ii_Byte
            End If
            tSTR = Trim(tSTR)
        End If
    End With
    lblOctet08.Caption = tSTR
    
End Sub
'*                                                                                                  *

Private Sub txtSI16_GotFocus()
    Call Select_TextBox_Str(txtSI16)
End Sub
'*                                                                                                  *

Private Sub txtSI16_KeyPress(KeyAscii As Integer)
    KeyAscii = Check_IntOnly(KeyAscii)
End Sub
'*                                                                                                  *

Private Sub txtSI16_KeyUp(KeyCode As Integer, Shift As Integer)
        
    Dim tSTR As String
    
    With txtSI16
        If .Text = "-" Then Exit Sub
        If IsNumeric(.Text) = False Then .Text = "0"
        If Val(.Text) < -32768 Then .Text = -32768
        If Val(.Text) > 32767 Then .Text = 32767
        tSTR = Replace(Format(Hex(CInt(.Text)), "@@@@"), " ", "0")
    End With
    lblSI16.Caption = Mid(tSTR, 1, 2) & " " & Mid(tSTR, 3, 2)
    
End Sub
'*                                                                                                  *

Private Sub txtSI16_LostFocus()
    With txtSI16
        If IsNumeric(.Text) = False Then
            .Text = "0"
            lblSI16.Caption = "00 00"
        End If
    End With
End Sub
'*                                                                                                  *

Private Sub txtUI16_GotFocus()
    Call Select_TextBox_Str(txtUI16)
End Sub
'*                                                                                                  *

Private Sub txtUI16_KeyPress(KeyAscii As Integer)
    KeyAscii = Check_UnsignedOnly(KeyAscii)
End Sub
'*                                                                                                  *

Private Sub txtUI16_KeyUp(KeyCode As Integer, Shift As Integer)
    
    Dim tSTR As String
    
    With txtUI16
        If IsNumeric(.Text) = False Then .Text = "0"
        If Val(.Text) < 0 Then .Text = 0
        If Val(.Text) > 65535 Then .Text = 65535
        tSTR = Replace(Format(Hex(CLng(.Text)), "@@@@"), " ", "0")
    End With
    lblUI16.Caption = Mid(tSTR, 1, 2) & " " & Mid(tSTR, 3, 2)

End Sub
'*                                                                                                  *

Private Sub txtUL32_GotFocus()
    Call Select_TextBox_Str(txtUL32)
End Sub
'*                                                                                                  *

Private Sub txtUL32_KeyPress(KeyAscii As Integer)
    KeyAscii = Check_UnsignedOnly(KeyAscii)
End Sub
'*                                                                                                  *

Private Sub txtUL32_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim tSTR As String
    Dim tDBL As Double
    
    With txtUL32
        If IsNumeric(.Text) = False Then .Text = "0"
        If Val(.Text) < 0 Then .Text = 0
        If Val(.Text) > 4294967295# Then .Text = 4294967295#
        If Val(.Text) >= &H1000000 Then
            tDBL = Int(Val(.Text) / &H10000)
            tSTR = Replace(Format(Hex(tDBL), "@@@@"), " ", "0")
            tSTR = tSTR & Replace(Format(Hex(Val(.Text) - (tDBL * &H10000)), "@@@@"), " ", "0")
        Else
            tSTR = Replace(Format(Hex(CDbl(.Text)), "@@@@@@@@"), " ", "0")
        End If
    End With
    lblUL32.Caption = Mid(tSTR, 1, 2) & " " & Mid(tSTR, 3, 2) & " " & _
                    Mid(tSTR, 5, 2) & " " & Mid(tSTR, 7, 2)

End Sub
'*                                                                                                  *

Private Sub txtOctet04_GotFocus()
    Call Select_TextBox_Str(txtOctet04)
End Sub
'*                                                                                                  *

Private Sub txtOctet04_KeyPress(KeyAscii As Integer)
    KeyAscii = Check_FloatOnly(KeyAscii)
End Sub
'*                                                                                                  *

Private Sub txtOctet04_KeyUp(KeyCode As Integer, Shift As Integer)
    
    Dim tSNG As Single
    Dim tBYTE(3) As Byte
    
    With txtOctet04
        If .Text = "-" Then Exit Sub
        If IsNumeric(.Text) = False Then .Text = "0"
        tSNG = CSng(.Text)
    End With
    CopyMemory tBYTE(0), tSNG, 4
    lblOctet04.Caption = HexToTwo(Hex(tBYTE(3))) & " " & HexToTwo(Hex(tBYTE(2))) & " " & _
                        HexToTwo(Hex(tBYTE(1))) & " " & HexToTwo(Hex(tBYTE(0)))

End Sub
'*                                                                                                  *

Private Sub txtOctet04_LostFocus()
    With txtOctet04
        If IsNumeric(.Text) = False Then
            .Text = "0"
            lblOctet04.Caption = "00 00 00 00"
        End If
    End With
End Sub
'*                                                                                                  *

Private Sub txtSB08_GotFocus()
    Call Select_TextBox_Str(txtSB08)
End Sub
'*                                                                                                  *

Private Sub txtSB08_KeyPress(KeyAscii As Integer)
    KeyAscii = Check_IntOnly(KeyAscii)
End Sub
'*                                                                                                  *

Private Sub txtSB08_KeyUp(KeyCode As Integer, Shift As Integer)
    
    With txtSB08
        If .Text = "-" Then Exit Sub
        If IsNumeric(.Text) = False Then .Text = "0"
        If Val(.Text) < -128 Then .Text = -128
        If Val(.Text) > 127 Then .Text = 127
        lblSB08.Caption = Right(Replace(Format(Hex(CByte(.Text)), "@@@@"), " ", "0"), 2)
    End With
    
End Sub
'*                                                                                                  *

Private Sub txtSB08_LostFocus()
    With txtSB08
        If IsNumeric(.Text) = False Then
            .Text = "0"
            lblSB08.Caption = "00"
        End If
    End With
End Sub
'*                                                                                                  *

Private Sub txtUB08_GotFocus()
    Call Select_TextBox_Str(txtUB08)
End Sub
'*                                                                                                  *

Private Sub txtUB08_KeyPress(KeyAscii As Integer)
    KeyAscii = Check_UnsignedOnly(KeyAscii)
End Sub
'*                                                                                                  *

Private Sub txtUB08_KeyUp(KeyCode As Integer, Shift As Integer)
    With txtUB08
        If IsNumeric(.Text) = False Then .Text = "0"
        If Val(.Text) < 0 Then .Text = 0
        If Val(.Text) > 255 Then .Text = 255
        lblUB08.Caption = HexToTwo(Hex(CByte(.Text)))
    End With
End Sub
'*                                                                                                  *

Private Sub txtOctet12Date_GotFocus(Index As Integer)
    Call Select_TextBox_Str(txtOctet12Date(Index))
End Sub
'*                                                                                                  *

Private Sub txtOctet12Date_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = Check_UnsignedOnly(KeyAscii)
End Sub
'*                                                                                                  *

Private Sub txtOctet12Date_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    Dim iYear As Integer, iMonth As Integer, iDay As Integer, bWeekDay As Byte
    Dim tSTR As String
    
    With txtOctet12Date(Index)
        If .Text = "" Then Exit Sub
        Select Case Index
            Case 0
                If IsNumeric(.Text) = False Then .Text = Year(Now)
                If Val(.Text) < 1 Then .Text = 1
                If Val(.Text) > 9999 Then .Text = 9999
            Case 1
                If IsNumeric(.Text) = False Then .Text = Month(Now)
                If Val(.Text) < 1 Then .Text = 1
                If Val(.Text) > 12 Then .Text = 12
            Case 2
                If IsNumeric(.Text) = False Then .Text = Day(Now)
                If Val(.Text) < 1 Then .Text = 1
                iYear = CInt(txtOctet12Date(0).Text):   iMonth = CInt(txtOctet12Date(1).Text)
                iDay = Find_Last_Day(iYear, iMonth)
                If Val(.Text) > iDay Then .Text = iDay
        End Select
    End With
    
    iYear = CInt(txtOctet12Date(0).Text)
    iMonth = CInt(txtOctet12Date(1).Text)
    iDay = CInt(txtOctet12Date(2).Text)
    bWeekDay = Weekday(DateSerial(iYear, iMonth, iDay), vbMonday)
    
    tSTR = Replace(Format(Hex(iYear), "@@@@"), " ", "0")
    tSTR = Mid(tSTR, 1, 2) & " " & Mid(tSTR, 3, 2) & " " & HexToTwo(Hex(iMonth)) & " " & _
            HexToTwo(Hex(iDay)) & " " & HexToTwo(Hex(bWeekDay))
    lblOctet12Date.Caption = tSTR

End Sub
'*                                                                                                  *

Private Sub txtOctet12Date_LostFocus(Index As Integer)
    With txtOctet12Date(Index)
        If .Text = "" Then
            Select Case Index
                Case 0:     .Text = Year(Now)
                Case 1:     .Text = Month(Now)
                Case 2:     .Text = Day(Now)
            End Select
        End If
    End With
    Call txtOctet12Date_KeyUp(0, 0, 0)
End Sub
'*                                                                                                  *

Private Sub txtOctet12Time_GotFocus(Index As Integer)
    Call Select_TextBox_Str(txtOctet12Time(Index))
End Sub
'*                                                                                                  *

Private Sub txtOctet12Time_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = Check_UnsignedOnly(KeyAscii)
End Sub
'*                                                                                                  *

Private Sub txtOctet12Time_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    Dim bHour As Byte, bMinute As Byte, bSecond As Byte
    Dim tSTR As String
    
    With txtOctet12Time(Index)
        If .Text = "" Then Exit Sub
        Select Case Index
            Case 0
                If IsNumeric(.Text) = False Then .Text = Hour(Now)
                If Val(.Text) < 0 Then .Text = 0
                If Val(.Text) > 23 Then .Text = 23
            Case 1
                If IsNumeric(.Text) = False Then .Text = Minute(Now)
                If Val(.Text) < 0 Then .Text = 0
                If Val(.Text) > 59 Then .Text = 59
            Case 2
                If IsNumeric(.Text) = False Then .Text = Second(Now)
                If Val(.Text) < 0 Then .Text = 0
                If Val(.Text) > 59 Then .Text = 59
        End Select
    End With
    
    bHour = CInt(txtOctet12Time(0).Text)
    bMinute = CInt(txtOctet12Time(1).Text)
    bSecond = CInt(txtOctet12Time(2).Text)
    
    tSTR = HexToTwo(Hex(bHour)) & " " & HexToTwo(Hex(bMinute)) & " " & _
            HexToTwo(Hex(bSecond)) & " " & "FF 80 00 00"
    lblOctet12Time.Caption = tSTR

End Sub
'*                                                                                                  *

Private Sub txtOctet12Time_LostFocus(Index As Integer)
    With txtOctet12Time(Index)
        If .Text = "" Then
            Select Case Index
                Case 0:     .Text = Hour(Now)
                Case 1:     .Text = Minute(Now)
                Case 2:     .Text = Second(Now)
            End Select
        End If
    End With
    Call txtOctet12Time_KeyUp(0, 0, 0)
End Sub
'*                                                                                                  *

Private Sub tmrOctet12_Timer()
    
    Dim Now_Timer As Date
    
    Now_Timer = Now
    txtOctet12Date(0).Text = Year(Now_Timer):   txtOctet12Date(1).Text = Month(Now_Timer)
    txtOctet12Date(2).Text = Day(Now_Timer):    txtOctet12Time(0).Text = Hour(Now_Timer)
    txtOctet12Time(1).Text = Minute(Now_Timer): txtOctet12Time(2).Text = Second(Now_Timer)
    
    Call txtOctet12Date_KeyUp(0, 0, 0)
    Call txtOctet12Time_KeyUp(0, 0, 0)
    DoEvents
    
End Sub
'*                                                                                                  *

Private Sub chkOctet12_Click()
    tmrOctet12.Enabled = chkOctet12.Value
End Sub
'*                                                                                                  *


