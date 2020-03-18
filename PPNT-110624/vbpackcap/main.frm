VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "KS4600Utility"
   ClientHeight    =   11070
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   11070
   ScaleWidth      =   15240
   Begin VB.Frame Frame1 
      Caption         =   "Adaptors"
      Height          =   975
      Left            =   120
      TabIndex        =   74
      Top             =   120
      Width           =   10575
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   120
         TabIndex        =   76
         Top             =   300
         Width           =   9315
      End
      Begin VB.CommandButton cmdAdaptorsSet 
         Caption         =   "Set"
         Height          =   315
         Left            =   9600
         TabIndex        =   75
         Top             =   300
         Width           =   795
      End
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear List"
      Height          =   855
      Left            =   13440
      TabIndex        =   6
      Top             =   240
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   960
      Left            =   1200
      TabIndex        =   5
      Top             =   10680
      Visible         =   0   'False
      Width           =   12495
   End
   Begin VB.Frame Frame2 
      Caption         =   "Capture"
      Height          =   975
      Left            =   10800
      TabIndex        =   0
      Top             =   120
      Width           =   2415
      Begin VB.OptionButton Option2 
         Caption         =   "To disk"
         Height          =   255
         Left            =   180
         TabIndex        =   4
         Top             =   540
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "To memory"
         Height          =   195
         Left            =   180
         TabIndex        =   3
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.CommandButton cmdCaptureStop 
         Caption         =   "Stop"
         Height          =   315
         Left            =   1560
         TabIndex        =   2
         Top             =   540
         Width           =   675
      End
      Begin VB.CommandButton cmdCaptureStart 
         Caption         =   "Start"
         Enabled         =   0   'False
         Height          =   315
         Left            =   1560
         TabIndex        =   1
         Top             =   180
         Width           =   675
      End
   End
   Begin VB.Frame Frame3 
      Height          =   9285
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   15015
      Begin VB.PictureBox FreqScaler 
         Appearance      =   0  'Æò¸é
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  '¾øÀ½
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   7230
         ScaleHeight     =   375
         ScaleWidth      =   6975
         TabIndex        =   81
         Top             =   5760
         Width           =   6975
      End
      Begin VB.PictureBox BitScaler 
         Appearance      =   0  'Æò¸é
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  '¾øÀ½
         ForeColor       =   &H00800000&
         Height          =   4080
         Left            =   6810
         ScaleHeight     =   4080
         ScaleWidth      =   375
         TabIndex        =   80
         Top             =   1710
         Width           =   375
      End
      Begin VB.PictureBox PicDraw 
         AutoRedraw      =   -1  'True
         Height          =   4095
         Left            =   7200
         ScaleHeight     =   4035
         ScaleWidth      =   6915
         TabIndex        =   79
         Top             =   1680
         Width           =   6975
      End
      Begin VB.CommandButton cmdChIfo 
         Caption         =   "Channel Info."
         Height          =   1215
         Left            =   12615
         TabIndex        =   78
         Top             =   6885
         Width           =   1455
      End
      Begin VB.CommandButton cmdChEstimation 
         Caption         =   "Channel Estimation"
         Height          =   1215
         Left            =   10935
         TabIndex        =   77
         Top             =   6885
         Width           =   1455
      End
      Begin VB.CommandButton cmdSearchDevice 
         Caption         =   "SearchDevice"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3000
         TabIndex        =   12
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox txtMACAddress 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "00:00:00:00:00:00"
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox txtNumofDevice 
         Alignment       =   1  '¿À¸¥ÂÊ ¸ÂÃã
         Enabled         =   0   'False
         Height          =   270
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   720
         Width           =   735
      End
      Begin VB.ListBox ListSID 
         Height          =   420
         Left            =   720
         TabIndex        =   9
         Top             =   1200
         Width           =   1935
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   480
         Top             =   8160
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   120
         ImageHeight     =   130
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "main.frx":0000
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   2115
         Left            =   480
         TabIndex        =   13
         Top             =   2160
         Width           =   5520
         _ExtentX        =   9737
         _ExtentY        =   3731
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   1455
         Left            =   480
         TabIndex        =   14
         Top             =   4440
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   2566
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
         Height          =   990
         Left            =   480
         TabIndex        =   15
         Top             =   6000
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   1746
         _Version        =   393216
      End
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   1575
         Left            =   480
         TabIndex        =   8
         Top             =   7080
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   2778
         _Version        =   393217
         Style           =   7
         Appearance      =   1
      End
      Begin VB.Label Label12 
         Caption         =   "Mhz"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   14280
         TabIndex        =   83
         Top             =   5640
         Width           =   480
      End
      Begin VB.Label Label7 
         Caption         =   "Bit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6720
         TabIndex        =   82
         Top             =   1440
         Width           =   360
      End
      Begin VB.Label Label10 
         Caption         =   "MAC"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "SID"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label8 
         Caption         =   "Num of Device"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "PLC Information"
         Height          =   255
         Left            =   600
         TabIndex        =   16
         Top             =   1920
         Width           =   1695
      End
   End
   Begin VB.Frame Frame4 
      BorderStyle     =   0  '¾øÀ½
      Height          =   8775
      Left            =   120
      TabIndex        =   21
      Top             =   1680
      Width           =   15120
      Begin VB.CommandButton cmdSetting 
         Caption         =   "Default Change"
         Height          =   855
         Left            =   13200
         TabIndex        =   73
         Top             =   7800
         Width           =   1215
      End
      Begin VB.CommandButton cmdModify 
         Caption         =   "Modify"
         Height          =   855
         Left            =   11760
         TabIndex        =   72
         Top             =   7800
         Width           =   1215
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Refresh"
         Height          =   855
         Left            =   10320
         TabIndex        =   71
         Top             =   7800
         Width           =   1215
      End
      Begin VB.Frame Frame8 
         Caption         =   "XPLC 23 Parameter"
         Height          =   1455
         Left            =   6240
         TabIndex        =   62
         Top             =   6120
         Width           =   8775
         Begin VB.TextBox txtProgramImgUpNum 
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
            Left            =   7320
            TabIndex        =   69
            Text            =   "0"
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox txtSubCodeUpNum 
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
            Left            =   7320
            TabIndex        =   67
            Text            =   "0"
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox txtParentBPS 
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
            Left            =   2160
            TabIndex        =   65
            Text            =   "0"
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox txtParentSIDwrite 
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
            Left            =   1440
            TabIndex        =   63
            Text            =   "00:00:00:00:00:00"
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label34 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            Caption         =   "Program Image Update Num."
            Height          =   255
            Left            =   4080
            TabIndex        =   70
            Top             =   840
            Width           =   2775
         End
         Begin VB.Label Label33 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            Caption         =   "Sub-Code Update Num."
            Height          =   255
            Left            =   4200
            TabIndex        =   68
            Top             =   360
            Width           =   2055
         End
         Begin VB.Label Label32 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            Caption         =   "Parent BPS"
            Height          =   255
            Left            =   120
            TabIndex        =   66
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label31 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            Caption         =   "Parent SID"
            Height          =   255
            Left            =   120
            TabIndex        =   64
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.TextBox txtLinkStationId3 
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
         Left            =   8040
         TabIndex        =   60
         Text            =   "00:00:00:00:00:00"
         Top             =   5640
         Width           =   1815
      End
      Begin VB.TextBox txtLinkStationId2 
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
         Left            =   12120
         TabIndex        =   58
         Text            =   "00:00:00:00:00:00"
         Top             =   5040
         Width           =   1815
      End
      Begin VB.TextBox txtLinkStationId1 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BackColor       =   &H80000009&
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
         Left            =   8040
         TabIndex        =   57
         Text            =   "00:00:00:00:00:00"
         Top             =   5040
         Width           =   1815
      End
      Begin VB.TextBox txtUserDesc 
         Height          =   735
         Left            =   120
         TabIndex        =   55
         Top             =   6960
         Width           =   5415
      End
      Begin VB.Frame Frame5 
         Caption         =   "General Configuration"
         Height          =   2535
         Left            =   6480
         TabIndex        =   34
         Top             =   2400
         Width           =   7695
         Begin VB.Frame Frame7 
            Caption         =   "Station Mode"
            Height          =   975
            Left            =   3960
            TabIndex        =   49
            Top             =   1440
            Width           =   3495
            Begin VB.OptionButton OpStationModeProgramming 
               Caption         =   "Programming"
               Height          =   255
               Left            =   1800
               TabIndex        =   53
               Top             =   600
               Width           =   1575
            End
            Begin VB.OptionButton OpStationModeFactory 
               Caption         =   "Factory"
               Height          =   255
               Left            =   240
               TabIndex        =   52
               Top             =   600
               Width           =   1215
            End
            Begin VB.OptionButton OpStationModeSuspend 
               Caption         =   "Suspend"
               Height          =   255
               Left            =   1800
               TabIndex        =   51
               Top             =   240
               Width           =   1335
            End
            Begin VB.OptionButton OpStationModeActive 
               Caption         =   "Active"
               Height          =   255
               Left            =   240
               TabIndex        =   50
               Top             =   240
               Width           =   1335
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Bank Indicator"
            Height          =   615
            Left            =   3960
            TabIndex        =   46
            Top             =   720
            Width           =   3495
            Begin VB.OptionButton OpBankIndicator1 
               Caption         =   "Bank 1"
               Height          =   255
               Left            =   1800
               TabIndex        =   48
               Top             =   240
               Width           =   1095
            End
            Begin VB.OptionButton OpBankIndicator0 
               Caption         =   "Bank 0"
               Height          =   255
               Left            =   240
               TabIndex        =   47
               Top             =   240
               Width           =   1095
            End
         End
         Begin VB.TextBox txtMinwirte 
            Height          =   375
            Left            =   6360
            TabIndex        =   44
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox txtHourwrite 
            Height          =   375
            Left            =   5400
            TabIndex        =   42
            Top             =   240
            Width           =   375
         End
         Begin VB.CheckBox Check6 
            Caption         =   "Key Transmission Enable"
            Height          =   255
            Left            =   240
            TabIndex        =   40
            Top             =   2160
            Width           =   3375
         End
         Begin VB.CheckBox Check5 
            Caption         =   "Tx Notification Enable"
            Height          =   255
            Left            =   240
            TabIndex        =   39
            Top             =   1800
            Width           =   2655
         End
         Begin VB.CheckBox Check4 
            Caption         =   "RTS/CTS Enable"
            Height          =   255
            Left            =   240
            TabIndex        =   38
            Top             =   1440
            Width           =   3135
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Secondary Link Restriction Enable"
            Height          =   255
            Left            =   240
            TabIndex        =   37
            Top             =   1080
            Width           =   3375
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Link Restriction Enable"
            Height          =   255
            Left            =   240
            TabIndex        =   36
            Top             =   720
            Width           =   3495
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Secondary Interface Enable"
            Height          =   255
            Left            =   240
            TabIndex        =   35
            Top             =   360
            Width           =   3255
         End
         Begin VB.Label Label20 
            Caption         =   "Min"
            Height          =   255
            Left            =   6840
            TabIndex        =   45
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label19 
            Caption         =   "Hour"
            Height          =   255
            Left            =   5880
            TabIndex        =   43
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label18 
            Caption         =   "Reset Period"
            Height          =   255
            Left            =   4080
            TabIndex        =   41
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.TextBox txtFBBTTLwrite 
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
         Left            =   8280
         TabIndex        =   33
         Text            =   "00"
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox txt2ndwrite 
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
         Left            =   11640
         Locked          =   -1  'True
         TabIndex        =   31
         Text            =   "00:00:00:00:00:00:00"
         Top             =   1320
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
         Left            =   7200
         TabIndex        =   29
         Text            =   "00:00:00:00:00:00"
         Top             =   1320
         Width           =   1815
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
         Left            =   11640
         TabIndex        =   27
         Text            =   "00:00:00:00:00:00:00"
         Top             =   720
         Width           =   2055
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
         Left            =   7200
         TabIndex        =   25
         Text            =   "00:00:00:00:00:00"
         Top             =   720
         Width           =   1815
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGridParameter 
         Height          =   5505
         Left            =   105
         TabIndex        =   22
         Top             =   720
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   9710
         _Version        =   393216
      End
      Begin VB.Label Label30 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         Caption         =   "Link Station ID #3"
         Height          =   255
         Left            =   6480
         TabIndex        =   61
         Top             =   5760
         Width           =   1455
      End
      Begin VB.Label Label29 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         Caption         =   "Link Station ID #2"
         Height          =   255
         Left            =   10560
         TabIndex        =   59
         Top             =   5160
         Width           =   1455
      End
      Begin VB.Label Label28 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         Caption         =   "Link Station ID #1"
         Height          =   255
         Left            =   6480
         TabIndex        =   56
         Top             =   5160
         Width           =   1455
      End
      Begin VB.Label Label27 
         Caption         =   "User Description"
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   6600
         Width           =   1575
      End
      Begin VB.Label Label11 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         Caption         =   "FBB TTL"
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6480
         TabIndex        =   32
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label5 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         Caption         =   "2nd"
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10800
         TabIndex        =   30
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label4 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         Caption         =   "SID"
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6360
         TabIndex        =   28
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         Caption         =   "KEY"
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10800
         TabIndex        =   26
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         Caption         =   "GID"
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6360
         TabIndex        =   24
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Parameter Setting"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   1695
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   9735
      Left            =   0
      TabIndex        =   20
      Top             =   1320
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   17171
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*       Visual Basic Packet Capture Demo             *
'*              using vbpcap.dll                      *
'*                                                    *
'*           Alan Cordwell March 2005                 *
'*           vbpcap by Lorenzo Cerulli                *
'*                                                    *
'******************************************************

Option Explicit

Dim strHexDump As String
Dim bCapture As Boolean
Dim OwnMACAddress(7) As Byte
Dim cbuff(2000) As Byte
Dim SID_Node As Node
Dim fme As Integer
Dim DnVal(54) As Byte
Dim UpVal(54) As Byte


Const CURRENT = 1
Const DEFAULT = 2

Const UPSTREAM = 1
Const DOWNSTREAM = 2

Private Sub cmdChEstimation_Click()
    vpSetParam PRM_MODE, CAPTURE_PROMISCUOUS
    Channel_Estimaion_Req
End Sub

Private Sub cmdChIfo_Click()
    vpSetParam PRM_MODE, CAPTURE_PROMISCUOUS
    Channel_Info_Req
End Sub
Private Function Channel_Info_Req()
    Dim tmpByteArray(59) As Byte
    Dim SelSID(7) As Byte
    Dim SttID(7) As Byte
    'Channel_Estimaion_Req
    SelSID(0) = "&H" & Mid(ListSID.Text, 1, 2)
    SelSID(1) = "&H" & Mid(ListSID.Text, 4, 2)
    SelSID(2) = "&H" & Mid(ListSID.Text, 7, 2)
    SelSID(3) = "&H" & Mid(ListSID.Text, 10, 2)
    SelSID(4) = "&H" & Mid(ListSID.Text, 13, 2)
    SelSID(5) = "&H" & Mid(ListSID.Text, 16, 2)
    
    SttID(0) = "&H" & Mid(SID_Node, 1, 2)
    SttID(1) = "&H" & Mid(SID_Node, 4, 2)
    SttID(2) = "&H" & Mid(SID_Node, 7, 2)
    SttID(3) = "&H" & Mid(SID_Node, 10, 2)
    SttID(4) = "&H" & Mid(SID_Node, 13, 2)
    SttID(5) = "&H" & Mid(SID_Node, 16, 2)
    
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
    tmpByteArray(16) = &H0: tmpByteArray(17) = &H1B
    tmpByteArray(18) = &H40: tmpByteArray(19) = &HA
    tmpByteArray(20) = SttID(0): tmpByteArray(21) = SttID(1): tmpByteArray(22) = SttID(2)
    tmpByteArray(23) = SttID(3): tmpByteArray(24) = SttID(4): tmpByteArray(25) = SttID(5)
    tmpByteArray(26) = &H0: tmpByteArray(27) = &H0: tmpByteArray(28) = &H0: tmpByteArray(29) = &H0: tmpByteArray(30) = &H0
    tmpByteArray(31) = &H0: tmpByteArray(32) = &H0: tmpByteArray(33) = &H0: tmpByteArray(34) = &H0: tmpByteArray(35) = &H0
    tmpByteArray(36) = &H0: tmpByteArray(37) = &H0: tmpByteArray(38) = &H0: tmpByteArray(39) = &H0: tmpByteArray(40) = &H0
    tmpByteArray(41) = &H0: tmpByteArray(42) = &H0: tmpByteArray(43) = &H0: tmpByteArray(44) = &H0: tmpByteArray(45) = &H0
    tmpByteArray(46) = &H0: tmpByteArray(47) = &H0: tmpByteArray(48) = &H0: tmpByteArray(49) = &H0: tmpByteArray(50) = &H0
    tmpByteArray(51) = &H0: tmpByteArray(52) = &H0: tmpByteArray(53) = &H0: tmpByteArray(54) = &H0: tmpByteArray(55) = &H0
    tmpByteArray(56) = &H0: tmpByteArray(57) = &H0: tmpByteArray(58) = &H0: tmpByteArray(59) = &H0

    vpSendPacket tmpByteArray()
End Function

Private Sub cmdClear_Click()
List1.Clear
End Sub
Private Function Channel_Estimaion_Req()
    Dim tmpByteArray(59) As Byte
    Dim SelSID(7) As Byte
    Dim SttID(7) As Byte
    'Channel_Estimaion_Req
    SelSID(0) = "&H" & Mid(ListSID.Text, 1, 2)
    SelSID(1) = "&H" & Mid(ListSID.Text, 4, 2)
    SelSID(2) = "&H" & Mid(ListSID.Text, 7, 2)
    SelSID(3) = "&H" & Mid(ListSID.Text, 10, 2)
    SelSID(4) = "&H" & Mid(ListSID.Text, 13, 2)
    SelSID(5) = "&H" & Mid(ListSID.Text, 16, 2)
    
    SttID(0) = "&H" & Mid(SID_Node, 1, 2)
    SttID(1) = "&H" & Mid(SID_Node, 4, 2)
    SttID(2) = "&H" & Mid(SID_Node, 7, 2)
    SttID(3) = "&H" & Mid(SID_Node, 10, 2)
    SttID(4) = "&H" & Mid(SID_Node, 13, 2)
    SttID(5) = "&H" & Mid(SID_Node, 16, 2)
    
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
    tmpByteArray(16) = &H0: tmpByteArray(17) = &H4E
    tmpByteArray(18) = &H40: tmpByteArray(19) = &H52
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

Private Function Search_Device_Req()
    Dim tmpByteArray(59) As Byte
    'Search_Device_Req
    tmpByteArray(0) = &HFF: tmpByteArray(1) = &HFF: tmpByteArray(2) = &HFF 'promiscuous?
    tmpByteArray(3) = &HFF: tmpByteArray(4) = &HFF: tmpByteArray(5) = &HFF
    tmpByteArray(6) = OwnMACAddress(0) '&H0
    tmpByteArray(7) = OwnMACAddress(1) '&H23
    tmpByteArray(8) = OwnMACAddress(2) '&H26
    tmpByteArray(9) = OwnMACAddress(3) '&H37
    tmpByteArray(10) = OwnMACAddress(4) '&H85
    tmpByteArray(11) = OwnMACAddress(5) '&H6E
    tmpByteArray(12) = &H88: tmpByteArray(13) = &HC9
    tmpByteArray(14) = &H0: tmpByteArray(15) = &H1
    tmpByteArray(16) = &H0: tmpByteArray(17) = &H1
    tmpByteArray(18) = &H41: tmpByteArray(19) = &H2
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

Private Function Get_Parameter_Req()
    Dim tmpByteArray(59) As Byte
    Dim SelSID(7) As Byte
    'Get_Parameter_Req
    SelSID(0) = "&H" & Mid(ListSID.Text, 1, 2)
    SelSID(1) = "&H" & Mid(ListSID.Text, 4, 2)
    SelSID(2) = "&H" & Mid(ListSID.Text, 7, 2)
    SelSID(3) = "&H" & Mid(ListSID.Text, 10, 2)
    SelSID(4) = "&H" & Mid(ListSID.Text, 13, 2)
    SelSID(5) = "&H" & Mid(ListSID.Text, 16, 2)
    
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
Private Function Modify_Req() '110602-jjang
    Dim tmpByteArray(179) As Byte
    Dim SelSID(7) As Byte
    Dim ParentBPS(1) As Byte
    'Modify_Req
    SelSID(0) = "&H" & Mid(ListSID.Text, 1, 2)
    SelSID(1) = "&H" & Mid(ListSID.Text, 4, 2)
    SelSID(2) = "&H" & Mid(ListSID.Text, 7, 2)
    SelSID(3) = "&H" & Mid(ListSID.Text, 10, 2)
    SelSID(4) = "&H" & Mid(ListSID.Text, 13, 2)
    SelSID(5) = "&H" & Mid(ListSID.Text, 16, 2)
    If Len(txtParentBPS.Text) = 1 Then
        txtParentBPS.Text = "000" & txtParentBPS.Text
    ElseIf Len(txtParentBPS.Text) = 2 Then
        txtParentBPS.Text = "00" & txtParentBPS.Text
    ElseIf Len(txtParentBPS.Text) = 3 Then
        txtParentBPS.Text = "0" & txtParentBPS.Text
    End If
    ParentBPS(0) = "&H" & Mid(txtParentBPS.Text, 1, 2)
    ParentBPS(1) = "&H" & Mid(txtParentBPS.Text, 3, 2)
    
    tmpByteArray(0) = SelSID(0): tmpByteArray(1) = SelSID(1): tmpByteArray(2) = SelSID(2)
    tmpByteArray(3) = SelSID(3): tmpByteArray(4) = SelSID(4): tmpByteArray(5) = SelSID(5)
    tmpByteArray(6) = OwnMACAddress(0) '&H0
    tmpByteArray(7) = OwnMACAddress(1) '&H00
    tmpByteArray(8) = OwnMACAddress(2) '&H07
    tmpByteArray(9) = OwnMACAddress(3) '&H7f
    tmpByteArray(10) = OwnMACAddress(4) '&Hff
    tmpByteArray(11) = OwnMACAddress(5) '&H02
    tmpByteArray(12) = &H88: tmpByteArray(13) = &HC9
    tmpByteArray(14) = &H0: tmpByteArray(15) = &H1
    tmpByteArray(16) = &H0: tmpByteArray(17) = &HA   'SeqNum
    tmpByteArray(18) = &H40: tmpByteArray(19) = &H12 'cmd
    tmpByteArray(20) = SelSID(0): tmpByteArray(21) = SelSID(1): tmpByteArray(22) = SelSID(2)
    tmpByteArray(23) = SelSID(3): tmpByteArray(24) = SelSID(4): tmpByteArray(25) = SelSID(5)
    tmpByteArray(26) = "&H" & Mid(txtKEYwrite.Text, 1, 2): tmpByteArray(27) = "&H" & Mid(txtKEYwrite.Text, 4, 2): tmpByteArray(28) = "&H" & Mid(txtKEYwrite.Text, 7, 2)
    tmpByteArray(29) = "&H" & Mid(txtKEYwrite.Text, 10, 2): tmpByteArray(30) = "&H" & Mid(txtKEYwrite.Text, 13, 2): tmpByteArray(31) = "&H" & Mid(txtKEYwrite.Text, 16, 2): tmpByteArray(32) = "&H" & Mid(txtKEYwrite.Text, 19, 2)
    tmpByteArray(33) = "&H" & Mid(txt2ndwrite.Text, 1, 2): tmpByteArray(34) = "&H" & Mid(txt2ndwrite.Text, 4, 2): tmpByteArray(35) = "&H" & Mid(txt2ndwrite.Text, 7, 2): tmpByteArray(36) = "&H" & Mid(txt2ndwrite.Text, 10, 2): tmpByteArray(37) = "&H" & Mid(txt2ndwrite.Text, 13, 2)
    tmpByteArray(38) = "&H" & Mid(txt2ndwrite.Text, 16, 2): tmpByteArray(39) = "&H" & Mid(txt2ndwrite.Text, 19, 2)
    tmpByteArray(40) = &H0: tmpByteArray(41) = &H60: tmpByteArray(42) = &H80: tmpByteArray(43) = &H83 'General configuration
    tmpByteArray(44) = &HF: tmpByteArray(45) = &H0 'Serial setup
    tmpByteArray(46) = "&H" & Mid(txtLinkStationId1.Text, 1, 2): tmpByteArray(47) = "&H" & Mid(txtLinkStationId1.Text, 4, 2): tmpByteArray(48) = "&H" & Mid(txtLinkStationId1.Text, 7, 2)    'LinkStationID1
    tmpByteArray(49) = "&H" & Mid(txtLinkStationId1.Text, 10, 2): tmpByteArray(50) = "&H" & Mid(txtLinkStationId1.Text, 13, 2): tmpByteArray(51) = "&H" & Mid(txtLinkStationId1.Text, 16, 2) 'LinkStationID1
    tmpByteArray(52) = "&H" & Mid(txtLinkStationId2.Text, 1, 2): tmpByteArray(53) = "&H" & Mid(txtLinkStationId2.Text, 4, 2): tmpByteArray(54) = "&H" & Mid(txtLinkStationId2.Text, 7, 2)    'LinkStationID2
    tmpByteArray(55) = "&H" & Mid(txtLinkStationId2.Text, 10, 2): tmpByteArray(56) = "&H" & Mid(txtLinkStationId2.Text, 13, 2): tmpByteArray(57) = "&H" & Mid(txtLinkStationId2.Text, 16, 2) 'LinkStationID2
    tmpByteArray(58) = "&H" & Mid(txtLinkStationId3.Text, 1, 2): tmpByteArray(59) = "&H" & Mid(txtLinkStationId3.Text, 4, 2): tmpByteArray(60) = "&H" & Mid(txtLinkStationId3.Text, 7, 2)    'LinkStationID3
    tmpByteArray(61) = "&H" & Mid(txtLinkStationId3.Text, 10, 2): tmpByteArray(62) = "&H" & Mid(txtLinkStationId3.Text, 13, 2): tmpByteArray(63) = "&H" & Mid(txtLinkStationId1.Text, 16, 2) 'LinkStationID3
    tmpByteArray(64) = "&H" & Mid(txtParentSIDwrite.Text, 1, 2): tmpByteArray(65) = "&H" & Mid(txtParentSIDwrite.Text, 4, 2): tmpByteArray(66) = "&H" & Mid(txtParentSIDwrite.Text, 7, 2)    'LinkStationID3
    tmpByteArray(67) = "&H" & Mid(txtParentSIDwrite.Text, 10, 2): tmpByteArray(68) = "&H" & Mid(txtParentSIDwrite.Text, 13, 2): tmpByteArray(69) = "&H" & Mid(txtParentSIDwrite.Text, 16, 2) 'LinkStationID3
    tmpByteArray(70) = ParentBPS(0): tmpByteArray(71) = ParentBPS(1) 'ParentBPS
    tmpByteArray(72) = &H0: tmpByteArray(73) = &H0 'Sub-Code Update Num?
    tmpByteArray(74) = &H0: tmpByteArray(75) = &H0 'Program Image Update Num?
    tmpByteArray(76) = "&H" & Mid(txtGIDwrite.Text, 1, 2): tmpByteArray(77) = "&H" & Mid(txtGIDwrite.Text, 4, 2): tmpByteArray(78) = "&H" & Mid(txtGIDwrite.Text, 7, 2)    'GID
    tmpByteArray(79) = "&H" & Mid(txtGIDwrite.Text, 10, 2): tmpByteArray(80) = "&H" & Mid(txtGIDwrite.Text, 13, 2): tmpByteArray(81) = "&H" & Mid(txtGIDwrite.Text, 16, 2) 'GID
    tmpByteArray(82) = "&H" & Mid(txtSIDwrite.Text, 1, 2): tmpByteArray(83) = "&H" & Mid(txtSIDwrite.Text, 4, 2): tmpByteArray(84) = "&H" & Mid(txtSIDwrite.Text, 7, 2)    'SID
    tmpByteArray(85) = "&H" & Mid(txtSIDwrite.Text, 10, 2): tmpByteArray(86) = "&H" & Mid(txtSIDwrite.Text, 13, 2): tmpByteArray(87) = "&H" & Mid(txtSIDwrite.Text, 16, 2) 'SID
    tmpByteArray(88) = &H0: tmpByteArray(89) = &H0: tmpByteArray(90) = &H0: tmpByteArray(91) = &H0: tmpByteArray(92) = &H0: tmpByteArray(93) = &H0
    tmpByteArray(94) = &H0: tmpByteArray(95) = &H0: tmpByteArray(96) = &H0: tmpByteArray(97) = &H0: tmpByteArray(98) = &H0: tmpByteArray(99) = &H0
    tmpByteArray(100) = &H0: tmpByteArray(101) = &H0: tmpByteArray(102) = &H0: tmpByteArray(103) = &H0: tmpByteArray(104) = &H0: tmpByteArray(105) = &H0
    tmpByteArray(106) = &H0: tmpByteArray(107) = &H0: tmpByteArray(108) = &H0: tmpByteArray(109) = &H0: tmpByteArray(110) = &H0: tmpByteArray(111) = &H0
    tmpByteArray(112) = &H0: tmpByteArray(113) = &H0: tmpByteArray(114) = &H0: tmpByteArray(115) = &H0
    'User Description
    tmpByteArray(116) = &H0: tmpByteArray(117) = &H0: tmpByteArray(118) = &H0: tmpByteArray(119) = &H0: tmpByteArray(120) = &H0: tmpByteArray(121) = &H0
    tmpByteArray(122) = &H0: tmpByteArray(123) = &H0: tmpByteArray(124) = &H0: tmpByteArray(125) = &H0: tmpByteArray(126) = &H0: tmpByteArray(127) = &H0
    tmpByteArray(128) = &H0: tmpByteArray(129) = &H0: tmpByteArray(130) = &H0: tmpByteArray(131) = &H0: tmpByteArray(132) = &H0: tmpByteArray(133) = &H0
    tmpByteArray(134) = &H0: tmpByteArray(135) = &H0: tmpByteArray(136) = &H0: tmpByteArray(137) = &H0: tmpByteArray(138) = &H0: tmpByteArray(139) = &H0
    tmpByteArray(140) = &H0: tmpByteArray(141) = &H0: tmpByteArray(142) = &H0: tmpByteArray(143) = &H0: tmpByteArray(144) = &H0: tmpByteArray(145) = &H0
    tmpByteArray(146) = &H0: tmpByteArray(147) = &H0: tmpByteArray(148) = &H0: tmpByteArray(149) = &H0: tmpByteArray(150) = &H0: tmpByteArray(151) = &H0
    tmpByteArray(152) = &H0: tmpByteArray(153) = &H0: tmpByteArray(154) = &H0: tmpByteArray(155) = &H0: tmpByteArray(156) = &H0: tmpByteArray(157) = &H0
    tmpByteArray(158) = &H0: tmpByteArray(159) = &H0: tmpByteArray(160) = &H0: tmpByteArray(161) = &H0: tmpByteArray(162) = &H0: tmpByteArray(163) = &H0
    tmpByteArray(164) = &H0: tmpByteArray(165) = &H0: tmpByteArray(166) = &H0: tmpByteArray(167) = &H0: tmpByteArray(168) = &H0: tmpByteArray(169) = &H0
    tmpByteArray(170) = &H0: tmpByteArray(171) = &H0: tmpByteArray(172) = &H0: tmpByteArray(173) = &H0: tmpByteArray(174) = &H0: tmpByteArray(175) = &H0
    tmpByteArray(176) = &H0: tmpByteArray(177) = &H0: tmpByteArray(178) = &H0: tmpByteArray(179) = &H0
    
    vpSendPacket tmpByteArray()

End Function

Private Function Get_PacketInfo_Req()
    Dim tmpByteArray(59) As Byte
    Dim SelSID(7) As Byte
    'Get_PacketInfo_Req
    SelSID(0) = "&H" & Mid(ListSID.Text, 1, 2)
    SelSID(1) = "&H" & Mid(ListSID.Text, 4, 2)
    SelSID(2) = "&H" & Mid(ListSID.Text, 7, 2)
    SelSID(3) = "&H" & Mid(ListSID.Text, 10, 2)
    SelSID(4) = "&H" & Mid(ListSID.Text, 13, 2)
    SelSID(5) = "&H" & Mid(ListSID.Text, 16, 2)
    
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
    tmpByteArray(16) = &H0: tmpByteArray(17) = &H3 'SeqNum
    tmpByteArray(18) = &H0: tmpByteArray(19) = &H26
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

Private Function Get_StreamInfo_Req()
    Dim tmpByteArray(59) As Byte
    Dim SelSID(7) As Byte
    'Get_StreamInfo_Req
    SelSID(0) = "&H" & Mid(ListSID.Text, 1, 2)
    SelSID(1) = "&H" & Mid(ListSID.Text, 4, 2)
    SelSID(2) = "&H" & Mid(ListSID.Text, 7, 2)
    SelSID(3) = "&H" & Mid(ListSID.Text, 10, 2)
    SelSID(4) = "&H" & Mid(ListSID.Text, 13, 2)
    SelSID(5) = "&H" & Mid(ListSID.Text, 16, 2)
    
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
    tmpByteArray(16) = &H0: tmpByteArray(17) = &H38 'SeqNum
    tmpByteArray(18) = &H1: tmpByteArray(19) = &H46 '110603-jjang
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

Private Sub cmdModify_Click()
    vpSetParam PRM_MODE, CAPTURE_PROMISCUOUS
    Modify_Req
End Sub

Private Sub cmdSearchDevice_Click()
    ListSID.Clear
    vpSetParam PRM_MODE, CAPTURE_PROMISCUOUS
    Search_Device_Req
End Sub
Private Sub cmdSetting_Click()
Load Form2
Form2.Show vbModeless
   
End Sub

Private Sub ListSID_Click()
If ListSID.ListCount > 0 Then

    Get_Parameter_Req
    TreeView1.Nodes.Clear
    Set SID_Node = TreeView1.Nodes.Add(, , "Root", ListSID.Text)
End If
End Sub

Private Sub cmdAdaptorsSet_Click()
   Dim lTemp As Long
   
   'set the network adapter to that selected in combo1
   vpSetCurrentAdapter Combo1.ListIndex
   
   'check to see that the adapter has been set
   lTemp = vpGetCurrentAdapter
   If lTemp = Combo1.ListIndex Then
      'if so, enable the start button
      cmdCaptureStart.Enabled = True
   Else
      'if not, give error message
      MsgBox "Failed to set network adapter for capture.", vbExclamation
   End If
   
End Sub

Private Sub cmdCaptureStart_Click()
   'start button
   cmdSearchDevice.Enabled = True
   bCapture = True
   Call doCapture
End Sub

Private Sub cmdCaptureStop_Click()
  'stop button
  bCapture = False
End Sub

Private Function GetMACAddress() As String
    Dim obj, objs
    
    Set objs = GetObject("winmgmts:").ExecQuery("SELECT MACAddress FROM Win32_NetworkAdapter WHERE MACAddress Is Not NULL")

    For Each obj In objs
        GetMACAddress = obj.MACAddress
        Exit For
    Next obj
End Function

Private Function InitList()
    MSFlexGridParameter.Cols = 3
    MSFlexGridParameter.Rows = 24
    MSFlexGridParameter.ColWidth(0) = 2300
    MSFlexGridParameter.ColWidth(1) = 1800 'Current
    MSFlexGridParameter.ColWidth(2) = 1800 'Default
  
    MSFlexGridParameter.TextMatrix(0, 0) = "Parameter Name"
    MSFlexGridParameter.TextMatrix(0, 1) = "Current setting"
    MSFlexGridParameter.TextMatrix(0, 2) = "Default setting"
    
    MSFlexGridParameter.TextMatrix(1, 0) = "Program Version"
    MSFlexGridParameter.TextMatrix(2, 0) = "Sub-code Version"
    MSFlexGridParameter.TextMatrix(3, 0) = "Station ID"
    MSFlexGridParameter.TextMatrix(4, 0) = "Group ID"
    MSFlexGridParameter.TextMatrix(5, 0) = "Device Type"
    MSFlexGridParameter.TextMatrix(6, 0) = "Operation Mode"
    MSFlexGridParameter.TextMatrix(7, 0) = "Serial Parameter"
    MSFlexGridParameter.TextMatrix(8, 0) = "Self Reset Period"
    MSFlexGridParameter.TextMatrix(9, 0) = "2nd Interface Enable"
    MSFlexGridParameter.TextMatrix(10, 0) = "RTS CTS Status"
    MSFlexGridParameter.TextMatrix(11, 0) = "Link Restriction"
    MSFlexGridParameter.TextMatrix(12, 0) = "2nd Link Restriction"
    MSFlexGridParameter.TextMatrix(13, 0) = "Tx Notification"
    MSFlexGridParameter.TextMatrix(14, 0) = "Key Transmission"
    MSFlexGridParameter.TextMatrix(15, 0) = "Cumulate MSDU number"
    MSFlexGridParameter.TextMatrix(16, 0) = "FBB TTL number"
    MSFlexGridParameter.TextMatrix(17, 0) = "Firmware Upgrade Info."
    MSFlexGridParameter.TextMatrix(18, 0) = "Sub-Code Upgrade Info."
    MSFlexGridParameter.TextMatrix(19, 0) = "Tx FBB Status"
    MSFlexGridParameter.TextMatrix(20, 0) = "Tx Filter enable"
    MSFlexGridParameter.TextMatrix(21, 0) = "Parent Station ID"
    MSFlexGridParameter.TextMatrix(22, 0) = "Parent BPS"
    MSFlexGridParameter.TextMatrix(23, 0) = "Encryption Key"
    
    MSFlexGrid1.Cols = 3
    MSFlexGrid1.Rows = 9
    MSFlexGrid1.ColWidth(0) = 2800
    MSFlexGrid1.ColWidth(1) = 1300
    MSFlexGrid1.ColWidth(2) = 1300
  
    MSFlexGrid1.TextMatrix(0, 0) = "Item Name"
    MSFlexGrid1.TextMatrix(0, 1) = "Up Stream"
    MSFlexGrid1.TextMatrix(0, 2) = "Down Stream"
    
    MSFlexGrid1.TextMatrix(1, 0) = "Number of ACK Counter"
    MSFlexGrid1.TextMatrix(2, 0) = "Number of FAIL Counter"
    MSFlexGrid1.TextMatrix(3, 0) = "Number of Ethernet Bytes"
    MSFlexGrid1.TextMatrix(4, 0) = "Number of Ethernet Packets"
    MSFlexGrid1.TextMatrix(5, 0) = "Number of Serial Bytes"
    MSFlexGrid1.TextMatrix(6, 0) = "Number of Serial Packets"
    MSFlexGrid1.TextMatrix(7, 0) = "Parent Link Channel Information"
    MSFlexGrid1.TextMatrix(8, 0) = "Path Link Channel Information"

    MSFlexGrid2.Cols = 2
    MSFlexGrid2.Rows = 6
    MSFlexGrid2.ColWidth(0) = 4200
    MSFlexGrid2.ColWidth(1) = 1200
  
    MSFlexGrid2.TextMatrix(0, 0) = "Item Name"
    MSFlexGrid2.TextMatrix(0, 1) = "Value"

    MSFlexGrid2.TextMatrix(1, 0) = "Number of Transmitted Packets no Response"
    MSFlexGrid2.TextMatrix(2, 0) = "Number of Discarded Ethernet Packets"
    MSFlexGrid2.TextMatrix(3, 0) = "Number of Received packets CFCS Error"
    MSFlexGrid2.TextMatrix(4, 0) = "Number of Received packets DFCS Error"
    MSFlexGrid2.TextMatrix(5, 0) = "Number of Active Links"
    
    MSFlexGrid3.Cols = 3
    MSFlexGrid3.Rows = 4
    MSFlexGrid3.ColWidth(0) = 1500
    MSFlexGrid3.ColWidth(1) = 1300
    MSFlexGrid3.ColWidth(2) = 1300

    MSFlexGrid3.TextMatrix(0, 0) = "Item Name"
    MSFlexGrid3.TextMatrix(0, 1) = "Up Stream"
    MSFlexGrid3.TextMatrix(0, 2) = "Down Stream"

    MSFlexGrid3.TextMatrix(1, 0) = "Bits/Symbol"
    MSFlexGrid3.TextMatrix(2, 0) = "AGC Gain"
    MSFlexGrid3.TextMatrix(3, 0) = "Puncturing Flag"

End Function

Private Sub Form_Load()
   Dim numadapters As Long
   Dim i As Long
   Dim name As String
   Dim desc As String
   

   
   txtMACAddress.Text = GetMACAddress
   
   OwnMACAddress(0) = "&H" & Mid(txtMACAddress.Text, 1, 2)
   OwnMACAddress(1) = "&H" & Mid(txtMACAddress.Text, 4, 2)
   OwnMACAddress(2) = "&H" & Mid(txtMACAddress.Text, 7, 2)
   OwnMACAddress(3) = "&H" & Mid(txtMACAddress.Text, 10, 2)
   OwnMACAddress(4) = "&H" & Mid(txtMACAddress.Text, 13, 2)
   OwnMACAddress(5) = "&H" & Mid(txtMACAddress.Text, 16, 2)
   
    DrawScaleGrid
    DrawGrid
    PicDraw.CurrentX = 0
    PicDraw.CurrentY = 3000
   
   
   InitList
   'TreeView1.ImageList = ImageList1
   TreeView1.Nodes.Clear

   'set path to current directory- so you can find any files created!!
   ChDir App.Path
   
   'Initialise VBPCap and enumerate adapters
   numadapters = VBPcapInit
   DoEvents
      
   
   Frame3.Visible = True
   Frame4.Visible = False
   
   TabStrip1.Tabs.Add , , "PLC Information"
   TabStrip1.Tabs.Add , , "Parameter Setting"
   
   Unload Form2
   'populate Combo1 with the adapters descriptions
   For i = 0 To numadapters - 1
      vpGetAdapterInfoVB5 i, name, desc
      Combo1.AddItem desc
      If InStr(desc, "Connect") > 0 Then
        Combo1.ListIndex = i
      End If
   Next i
   
   'select default adapter or else give error message
   If Combo1.ListCount > 0 Then
      'Combo1.ListIndex = 0
      cmdAdaptorsSet_Click
   Else
      MsgBox "No network adaptors found!"
   End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   'close down vbpcap and exit application
   vpEnd
   Unload Form1
   End
End Sub

Private Function Parse_Search_Device_Res()
   Dim i As Integer
   Dim strSID As String
    For i = 6 To 11
       If cbuff(i) < &H10 Then
           strSID = strSID & ":0" & Hex(cbuff(i))
       Else
           strSID = strSID & ":" & Hex(cbuff(i))
       End If
    Next i
    strSID = Mid(strSID, 2, Len(strSID) - 1)
    ListSID.AddItem strSID
    txtNumofDevice.Text = ListSID.ListCount
End Function
Private Sub DrawScaleGrid()
    Dim i As Integer
    
    For i = 0 To 40
        If (i Mod 5) = 0 Then
            BitScaler.Line (260, 4000 - (i * 100))-(380, 4000 - (i * 100)), vbBlack
        Else
            BitScaler.Line (330, 4000 - (i * 100))-(380, 4000 - (i * 100)), vbBlack
        End If
    Next
    For i = 0 To 100
        If (i Mod 5) = 0 Then
            FreqScaler.Line (i * 100, 0)-(i * 100, 120), vbBlack
        Else
            FreqScaler.Line (i * 100, 0)-(i * 100, 40), vbBlack
        End If
    Next
    
End Sub
Private Sub DrawGrid()
    Dim x, y As Integer
    
    PicDraw.Cls
    PicDraw.DrawWidth = 1
    PicDraw.DrawStyle = vbDot
    
    For x = 1 To 10
        PicDraw.Line (x * 1000, 0)-(x * 1000, 8000), vbBlack
    Next

    For y = 1 To 8
        PicDraw.Line (0, y * 1000)-(10000, y * 1000), vbBlack
    Next

End Sub

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
    MSFlexGridParameter.TextMatrix(1, CURRENT) = "Ver" & Str((cbuff(49) * &H100 + cbuff(48)) / 100)
    MSFlexGridParameter.TextMatrix(1, DEFAULT) = "Ver" & Str((cbuff(21) * &H100 + cbuff(20)) / 100)
    
    '"Sub-code Version"
    MSFlexGridParameter.TextMatrix(2, CURRENT) = "Ver" & Val((cbuff(51) * &H100 + cbuff(50)) / 1000)
    MSFlexGridParameter.TextMatrix(2, DEFAULT) = "Ver" & Val((cbuff(51) * &H100 + cbuff(50)) / 1000) '??
    
    '"Station ID"
    MSFlexGridParameter.TextMatrix(3, CURRENT) = GetId(54, 6): MSFlexGridParameter.TextMatrix(3, DEFAULT) = GetId(31, 6)
    txtSIDwrite.Text = GetId(54, 6)
    '"Group ID"
    MSFlexGridParameter.TextMatrix(4, CURRENT) = GetId(110, 6): MSFlexGridParameter.TextMatrix(4, DEFAULT) = GetId(110, 6)
    txtGIDwrite.Text = GetId(110, 6)
    '"Device Type"
    Select Case (cbuff(53) And &H3)
        Case 0: MSFlexGridParameter.TextMatrix(5, CURRENT) = "Default"
        Case 1: MSFlexGridParameter.TextMatrix(5, CURRENT) = "Master"
        Case 2: MSFlexGridParameter.TextMatrix(5, CURRENT) = "Slave"
        Case 3: MSFlexGridParameter.TextMatrix(5, CURRENT) = "MeterGateway"
    End Select

    Select Case (cbuff(37) And &H3)
        Case 0: MSFlexGridParameter.TextMatrix(5, DEFAULT) = "Default"
        Case 1: MSFlexGridParameter.TextMatrix(5, DEFAULT) = "Master"
        Case 2: MSFlexGridParameter.TextMatrix(5, DEFAULT) = "Slave"
        Case 3: MSFlexGridParameter.TextMatrix(5, DEFAULT) = "MeterGateway"
    End Select
    
    '"Operation Mode"
    Select Case ((cbuff(75) And &HC) / &H100)
        Case 0: MSFlexGridParameter.TextMatrix(6, CURRENT) = "Active Mode"
        Case 1: MSFlexGridParameter.TextMatrix(6, CURRENT) = "Suspend Mode"
        Case 2: MSFlexGridParameter.TextMatrix(6, CURRENT) = "Factory Mode"
        Case 3: MSFlexGridParameter.TextMatrix(6, CURRENT) = "Programming Mode"
    End Select

    Select Case ((cbuff(39) And &HC) / &H100)
        Case 0: MSFlexGridParameter.TextMatrix(6, DEFAULT) = "Active Mode"
        Case 1: MSFlexGridParameter.TextMatrix(6, DEFAULT) = "Suspend Mode"
        Case 2: MSFlexGridParameter.TextMatrix(6, DEFAULT) = "Factory Mode"
        Case 3: MSFlexGridParameter.TextMatrix(6, DEFAULT) = "Programming Mode"
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
    MSFlexGridParameter.TextMatrix(7, CURRENT) = strSerialSet

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
    MSFlexGridParameter.TextMatrix(7, DEFAULT) = strSerialSet
    
    '"Self Reset Period"
    dValue = cbuff(75) * &H100 + cbuff(74)
    MSFlexGridParameter.TextMatrix(8, CURRENT) = Val((dValue And &H3C0) / &H40) & ":" & Val(dValue And &H3F) & "(Hour:Min)"
    dValue = cbuff(39) * &H100 + cbuff(38)
    MSFlexGridParameter.TextMatrix(8, DEFAULT) = Val((dValue And &H3C0) / &H40) & ":" & Val(dValue And &H3F) & "(Hour:Min)"
    
    '"2nd Interface Enable"
    If (cbuff(76) And &H20) = 0 Then
        MSFlexGridParameter.TextMatrix(9, CURRENT) = "Disable"
    Else
        MSFlexGridParameter.TextMatrix(9, CURRENT) = "Enable"
    End If
    
    If (cbuff(40) And &H20) = 0 Then
        MSFlexGridParameter.TextMatrix(9, DEFAULT) = "Disable"
    Else
        MSFlexGridParameter.TextMatrix(9, DEFAULT) = "Enable"
    End If
    
    ' "RTS CTS Status"
    If (cbuff(76) And &H10) = 0 Then
        MSFlexGridParameter.TextMatrix(10, CURRENT) = "Disable"
    Else
        MSFlexGridParameter.TextMatrix(10, CURRENT) = "Enable"
    End If
    
    If (cbuff(40) And &H10) = 0 Then
        MSFlexGridParameter.TextMatrix(10, DEFAULT) = "Disable"
    Else
        MSFlexGridParameter.TextMatrix(10, DEFAULT) = "Enable"
    End If
    
    '"Link Restriction"
    If (cbuff(75) And &H10) = 0 Then
        MSFlexGridParameter.TextMatrix(11, CURRENT) = "Disable"
    Else
        MSFlexGridParameter.TextMatrix(11, CURRENT) = "Enable"
    End If
    
    If (cbuff(39) And &H10) = 0 Then
        MSFlexGridParameter.TextMatrix(11, DEFAULT) = "Disable"
    Else
        MSFlexGridParameter.TextMatrix(11, DEFAULT) = "Enable"
    End If
    
    '"2nd Link Restriction"
    If (cbuff(76) And &H40) = 0 Then
        MSFlexGridParameter.TextMatrix(12, CURRENT) = "Disable"
    Else
        MSFlexGridParameter.TextMatrix(12, CURRENT) = "Enable"
    End If
    
    If (cbuff(40) And &H40) = 0 Then
        MSFlexGridParameter.TextMatrix(12, DEFAULT) = "Disable"
    Else
        MSFlexGridParameter.TextMatrix(12, DEFAULT) = "Enable"
    End If
    
    '"Tx Notification"
    If (cbuff(76) And &H8) = 0 Then
        MSFlexGridParameter.TextMatrix(13, CURRENT) = "Disable"
    Else
        MSFlexGridParameter.TextMatrix(13, CURRENT) = "Enable"
    End If
    
    If (cbuff(40) And &H8) = 0 Then
        MSFlexGridParameter.TextMatrix(13, DEFAULT) = "Disable"
    Else
        MSFlexGridParameter.TextMatrix(13, DEFAULT) = "Enable"
    End If
    
    '"Key Transmission"
    If (cbuff(76) And &H1) = 0 Then
        MSFlexGridParameter.TextMatrix(14, CURRENT) = "Disable"
    Else
        MSFlexGridParameter.TextMatrix(14, CURRENT) = "Enable"
    End If
    
    If (cbuff(40) And &H1) = 0 Then
        MSFlexGridParameter.TextMatrix(14, DEFAULT) = "Disable"
    Else
        MSFlexGridParameter.TextMatrix(14, DEFAULT) = "Enable"
    End If
    
    '"Cumulate MSDU number"
    MSFlexGridParameter.TextMatrix(15, CURRENT) = Val((cbuff(75) And &H60) / &H20)
    MSFlexGridParameter.TextMatrix(15, DEFAULT) = Val((cbuff(39) And &H60) / &H20)
    
    '"FBB TTL number"
    dValue = (((cbuff(77) And &H1F) * &H100 + cbuff(76)) And &H1F80) / &H80
    MSFlexGridParameter.TextMatrix(16, CURRENT) = Val(dValue)
    txtFBBTTLwrite = Val(dValue)
    dValue = (((cbuff(41) And &H1F) * &H100 + cbuff(40)) And &H1F80) / &H80
    MSFlexGridParameter.TextMatrix(16, DEFAULT) = Val(dValue)
    
    '"Firmware Upgrade Info."
    MSFlexGridParameter.TextMatrix(17, CURRENT) = "": MSFlexGridParameter.TextMatrix(17, DEFAULT) = ""
    
    '"Sub-Code Upgrade Info."
    MSFlexGridParameter.TextMatrix(18, CURRENT) = "": MSFlexGridParameter.TextMatrix(18, DEFAULT) = ""
    
    '"Tx FBB Status"
    MSFlexGridParameter.TextMatrix(19, CURRENT) = "": MSFlexGridParameter.TextMatrix(19, DEFAULT) = ""
    
    '"Tx Filter enable"
    MSFlexGridParameter.TextMatrix(20, CURRENT) = "": MSFlexGridParameter.TextMatrix(20, DEFAULT) = ""
    
    '"Parent Station ID"
    MSFlexGridParameter.TextMatrix(21, CURRENT) = GetId(98, 6): MSFlexGridParameter.TextMatrix(21, DEFAULT) = GetId(98, 6)
    txtParentSIDwrite.Text = GetId(98, 6)
    '"Parent BPS"
    MSFlexGridParameter.TextMatrix(22, CURRENT) = Val(cbuff(105) * &H100 + cbuff(104))
    MSFlexGridParameter.TextMatrix(22, DEFAULT) = Val(cbuff(105) * &H100 + cbuff(104)) '??
    txtParentBPS.Text = Val(cbuff(105) * &H100 + cbuff(104))
    '"Encryption Key"
    MSFlexGridParameter.TextMatrix(23, CURRENT) = GetId(60, 7): MSFlexGridParameter.TextMatrix(23, DEFAULT) = GetId(24, 7)
    txtKEYwrite.Text = GetId(60, 7)
    
End Function

Private Function Parse_Get_PacketInfo_Res()
    Dim dValue As Double
    
    '"Number of ACK Counter"
    dValue = CLng(cbuff(23)) * &H1000000 + CLng(cbuff(22)) * &H10000 + CLng(cbuff(21)) * &H100 + cbuff(20)
    MSFlexGrid1.TextMatrix(1, UPSTREAM) = Val(dValue)
    dValue = cbuff(27) * &H1000000 + cbuff(26) * &H10000 + cbuff(25) * &H100 + cbuff(24)
    MSFlexGrid1.TextMatrix(1, DOWNSTREAM) = Val(dValue)

    'Number of FAIL Counter"
    dValue = CLng(cbuff(31)) * &H1000000 + CLng(cbuff(30)) * &H10000 + CLng(cbuff(29)) * &H100 + cbuff(28)
    MSFlexGrid1.TextMatrix(2, UPSTREAM) = Val(dValue)
    dValue = CLng(cbuff(35)) * &H1000000 + CLng(cbuff(34)) * &H10000 + CLng(cbuff(33)) * &H100 + cbuff(32)
    MSFlexGrid1.TextMatrix(2, DOWNSTREAM) = Val(dValue)

    '"Number of Ethernet Bytes"
    dValue = CLng(cbuff(39)) * &H1000000 + CLng(cbuff(38)) * &H10000 + CLng(cbuff(37)) * &H100 + cbuff(36)
    MSFlexGrid1.TextMatrix(3, UPSTREAM) = Val(dValue)
    dValue = CLng(cbuff(43)) * &H1000000 + CLng(cbuff(42)) * &H10000 + CLng(cbuff(41)) * &H100 + cbuff(40)
    MSFlexGrid1.TextMatrix(3, DOWNSTREAM) = Val(dValue)

    '"Number of Ethernet Packets"
    dValue = CLng(cbuff(47)) * &H1000000 + CLng(cbuff(46)) * &H10000 + CLng(cbuff(45)) * &H100 + cbuff(44)
    MSFlexGrid1.TextMatrix(4, UPSTREAM) = Val(dValue)
    dValue = CLng(cbuff(51)) * &H1000000 + CLng(cbuff(50)) * &H10000 + CLng(cbuff(49)) * &H100 + cbuff(48)
    MSFlexGrid1.TextMatrix(4, DOWNSTREAM) = Val(dValue)

    '"Number of Serial Bytes"
    dValue = CLng(cbuff(55)) * &H1000000 + CLng(cbuff(54)) * &H10000 + CLng(cbuff(53)) * &H100 + cbuff(52)
    MSFlexGrid1.TextMatrix(5, UPSTREAM) = Val(dValue)
    dValue = CLng(cbuff(59)) * &H1000000 + CLng(cbuff(58)) * &H10000 + CLng(cbuff(57)) * &H100 + cbuff(56)
    MSFlexGrid1.TextMatrix(5, DOWNSTREAM) = Val(dValue)

    '"Number of Serial Packets"
    dValue = CLng(cbuff(63)) * &H1000000 + CLng(cbuff(62)) * &H10000 + CLng(cbuff(61)) * &H100 + cbuff(60)
    MSFlexGrid1.TextMatrix(6, UPSTREAM) = Val(dValue)
    dValue = CLng(cbuff(67)) * &H1000000 + CLng(cbuff(66)) * &H10000 + CLng(cbuff(65)) * &H100 + cbuff(64)
    MSFlexGrid1.TextMatrix(6, DOWNSTREAM) = Val(dValue)

    '"Parent Link Channel Information"
    '"Path Link Channel Information"
    

    '"Number of Transmitted Packets no Response"
    dValue = CLng(cbuff(71)) * &H1000000 + CLng(cbuff(70)) * &H10000 + CLng(cbuff(69)) * &H100 + cbuff(68)
    MSFlexGrid2.TextMatrix(1, 1) = Val(dValue)
    '"Number of Discarded Ethernet Packets"
    dValue = CLng(cbuff(75)) * &H1000000 + CLng(cbuff(74)) * &H10000 + CLng(cbuff(73)) * &H100 + cbuff(72)
    MSFlexGrid2.TextMatrix(2, 1) = Val(dValue)
    '"Number of Received packets CFCS Error"
    dValue = CLng(cbuff(79)) * &H1000000 + CLng(cbuff(78)) * &H10000 + CLng(cbuff(77)) * &H100 + cbuff(76)
    MSFlexGrid2.TextMatrix(3, 1) = Val(dValue)
    '"Number of Received packets DFCS Error"
    dValue = CLng(cbuff(83)) * &H1000000 + CLng(cbuff(82)) * &H10000 + CLng(cbuff(81)) * &H100 + cbuff(80)
    MSFlexGrid2.TextMatrix(4, 1) = Val(dValue)
    '"Number of Active Links"
    dValue = CLng(cbuff(87)) * &H1000000 + CLng(cbuff(86)) * &H10000 + CLng(cbuff(85)) * &H100 + cbuff(84)
    MSFlexGrid2.TextMatrix(5, 1) = Val(dValue)
    
End Function
Private Function Channel_Estimation_Res()
Call ListSID_Click
End Function
Private Function Channel_Info_Res()

DnVal(0) = cbuff(26): DnVal(1) = cbuff(27): DnVal(2) = cbuff(28): DnVal(3) = cbuff(29): DnVal(4) = cbuff(30): DnVal(5) = cbuff(31)
DnVal(6) = cbuff(32): DnVal(7) = cbuff(33): DnVal(8) = cbuff(34): DnVal(9) = cbuff(35): DnVal(10) = cbuff(36): DnVal(11) = cbuff(37)
DnVal(12) = cbuff(38): DnVal(13) = cbuff(39): DnVal(14) = cbuff(40): DnVal(15) = cbuff(41): DnVal(16) = cbuff(42): DnVal(17) = cbuff(43)
DnVal(18) = cbuff(44): DnVal(19) = cbuff(45): DnVal(20) = cbuff(46): DnVal(21) = cbuff(47): DnVal(22) = cbuff(48): DnVal(23) = cbuff(49)
DnVal(24) = cbuff(50): DnVal(25) = cbuff(51): DnVal(26) = cbuff(52): DnVal(27) = cbuff(53): DnVal(28) = cbuff(54): DnVal(29) = cbuff(55)
DnVal(30) = cbuff(56): DnVal(31) = cbuff(57): DnVal(32) = cbuff(58): DnVal(33) = cbuff(59): DnVal(34) = cbuff(60): DnVal(35) = cbuff(61)
DnVal(36) = cbuff(62): DnVal(37) = cbuff(63): DnVal(38) = cbuff(64): DnVal(39) = cbuff(65): DnVal(40) = cbuff(66): DnVal(41) = cbuff(67)
DnVal(42) = cbuff(68): DnVal(43) = cbuff(69): DnVal(44) = cbuff(70): DnVal(45) = cbuff(71): DnVal(46) = cbuff(72): DnVal(47) = cbuff(73)
DnVal(48) = cbuff(74): DnVal(49) = cbuff(75): DnVal(50) = cbuff(76): DnVal(51) = cbuff(77): DnVal(52) = cbuff(78): DnVal(53) = cbuff(79)

UpVal(0) = cbuff(80): UpVal(1) = cbuff(81): UpVal(2) = cbuff(82): UpVal(3) = cbuff(83): UpVal(4) = cbuff(84): UpVal(5) = cbuff(85)
UpVal(6) = cbuff(86): UpVal(7) = cbuff(87): UpVal(8) = cbuff(88): UpVal(9) = cbuff(89): UpVal(10) = cbuff(90): UpVal(11) = cbuff(91)
UpVal(12) = cbuff(92): UpVal(13) = cbuff(93): UpVal(14) = cbuff(94): UpVal(15) = cbuff(95): UpVal(16) = cbuff(96): UpVal(17) = cbuff(97)
UpVal(18) = cbuff(98): UpVal(19) = cbuff(99): UpVal(20) = cbuff(100): UpVal(21) = cbuff(101): UpVal(22) = cbuff(102): UpVal(23) = cbuff(103)
UpVal(24) = cbuff(104): UpVal(25) = cbuff(105): UpVal(26) = cbuff(106): UpVal(27) = cbuff(107): UpVal(28) = cbuff(108): UpVal(29) = cbuff(109)
UpVal(30) = cbuff(110): UpVal(31) = cbuff(111): UpVal(32) = cbuff(112): UpVal(33) = cbuff(113): UpVal(34) = cbuff(114): UpVal(35) = cbuff(115)
UpVal(36) = cbuff(116): UpVal(37) = cbuff(117): UpVal(38) = cbuff(118): UpVal(39) = cbuff(119): UpVal(40) = cbuff(120): UpVal(41) = cbuff(121)
UpVal(42) = cbuff(122): UpVal(43) = cbuff(123): UpVal(44) = cbuff(124): UpVal(45) = cbuff(125): UpVal(46) = cbuff(126): UpVal(47) = cbuff(127)
UpVal(48) = cbuff(128): UpVal(49) = cbuff(129): UpVal(50) = cbuff(130): UpVal(51) = cbuff(131): UpVal(52) = cbuff(132): UpVal(53) = cbuff(133)
PicDraw.Cls
DrawScaleGrid
DrawGraph
End Function
Private Sub DrawGraph()
    Dim i As Integer
    Dim gXScale As Integer
    Dim gXYShift As Integer
    
    gXScale = 150   'Shall be x10
    gXYShift = 15     'For prevent overwrite the previous graph
    
    PicDraw.CurrentX = 0
    PicDraw.CurrentY = 4000
    PicDraw.DrawWidth = 1
    PicDraw.DrawStyle = vbSolid
    
    For i = 0 To 53
        PicDraw.Line (i * gXScale, 4000 - 10 * DnVal(i))-((i + 0.1) * gXScale, 4000 - 10 * DnVal(i + 1)), vbBlue
        PicDraw.Line ((i + 0.1) * gXScale, 4000 - 10 * DnVal(i + 1))-((i + 1) * gXScale, 4000 - 10 * DnVal(i + 1)), vbBlue
    Next
    
    PicDraw.CurrentX = 0
    PicDraw.CurrentY = 4000
    PicDraw.DrawWidth = 1
    PicDraw.DrawStyle = vbSolid
    
    For i = 0 To 53
        PicDraw.Line ((i * gXScale) + gXYShift, 4000 - 10 * UpVal(i))-(((i + 0.1) * gXScale) + gXYShift, 4000 - 10 * UpVal(i + 1)), vbRed
        PicDraw.Line (((i + 0.1) * gXScale) + gXYShift, 4000 - gXYShift - 10 * UpVal(i + 1))-(((i + 1) * gXScale) + gXYShift, 4000 - gXYShift - 10 * UpVal(i + 1)), vbRed
    Next


End Sub
Private Function Parse_Get_StreamInfo_Res()
    Dim dValue As Long
    Dim strName As String
'SID
    strName = GetId(20, 6)
   Set SID_Node = TreeView1.Nodes.Add("Root", tvwChild, "Child1", strName)

'"Bits/Symbol"
    dValue = ((cbuff(27) And &H1) * &H100 + cbuff(26))
    MSFlexGrid3.TextMatrix(1, UPSTREAM) = Val(dValue)
    dValue = ((cbuff(29) And &H1) * &H100 + cbuff(28))
    MSFlexGrid3.TextMatrix(1, DOWNSTREAM) = Val(dValue)
'"AGC Gain"
    dValue = ((cbuff(27) And &H7E)) / &H2
    MSFlexGrid3.TextMatrix(2, UPSTREAM) = Val(dValue)
    dValue = ((cbuff(29) And &H7E)) / &H2
    MSFlexGrid3.TextMatrix(2, DOWNSTREAM) = Val(dValue)
'"Puncturing Flag"
    If (cbuff(27) And &HF0) = 0 Then
        MSFlexGrid3.TextMatrix(3, UPSTREAM) = "Disable"
    Else
        MSFlexGrid3.TextMatrix(3, UPSTREAM) = "Enable"
    End If
    If (cbuff(29) And &HF0) = 0 Then
        MSFlexGrid3.TextMatrix(3, DOWNSTREAM) = "Disable"
    Else
        MSFlexGrid3.TextMatrix(3, DOWNSTREAM) = "Enable"
    End If
End Function

Private Sub doCapture()
   Dim pktHeader As PacketHeader
   Dim lReturn As Long
   'Dim cbuff(2000) As Byte
   'Dim strHexDump As String
   Dim i As Integer
   
   'set up selected capture method
   vpSetParam PRM_KERNELBUFFSIZE, KERNELBUFFSIZE.A_1_MegaBytes
   If Option1.value = True Then
      vpSetParam PRM_DUMPTYPE, DUMPTYPE.MEM
   Else
      vpSetParam PRM_DUMPTYPE, DUMPTYPE.DISK_SAFE
      vpSetParam PRM_FILENAME, "test.dat"
   End If
   DoEvents
   
   'start capturing with 20 mS timeout
   vpBegin 20
   
    vpSetParam PRM_MODE, CAPTURE_PROMISCUOUS
   
   Do While bCapture
      'capture a packet if there is one
      lReturn = vpCapture(cbuff(), pktHeader)
      DoEvents
      'if so, display it
      If lReturn > 0 And pktHeader.caplen > 59 And cbuff(12) = &H88 And cbuff(13) = &HC9 Then
         'success
         strHexDump = ":"
         For i = 1 To pktHeader.caplen - 1
            If cbuff(i) < &H10 Then
                strHexDump = strHexDump & "0" & Hex(cbuff(i))
            Else
                strHexDump = strHexDump & Hex(cbuff(i))
            End If
         Next i
         List1.AddItem strHexDump
         
         If pktHeader.caplen = 60 And cbuff(19) = &H22 Then 'Search_Device_Res
            Parse_Search_Device_Res
         ElseIf pktHeader.caplen = 134 And cbuff(19) = &H1E Then  'Channel_Info_Res
         '   Parse_Get_PacketInfo_Res
            If ListSID.ListCount > 0 Then
                Channel_Info_Res
            End If
         ElseIf pktHeader.caplen = 60 And cbuff(19) = &H5A Then 'Channel_Estimation_Res
         '   Parse_Get_PacketInfo_Res
            If ListSID.ListCount > 0 Then
                Channel_Estimation_Res
            End If
         ElseIf pktHeader.caplen = 280 And cbuff(19) = &H1E Then 'Get_Parameter_Res
            Parse_Get_Parameter_Res
            If ListSID.ListCount > 0 Then
                Get_PacketInfo_Req
            End If
         ElseIf pktHeader.caplen = 94 And cbuff(19) = &H3E Then 'Get_PacketInfo_Res
            Parse_Get_PacketInfo_Res
            If ListSID.ListCount > 0 Then
                Get_StreamInfo_Req
            End If
         ElseIf pktHeader.caplen = 196 And cbuff(19) = &H5E Then 'Get_StreamInfo_Res
            Parse_Get_StreamInfo_Res
         End If
      Else
         'no packet captured
      End If
   Loop


End Sub

Private Sub TabStrip1_Click()

If TabStrip1.SelectedItem.Index = 1 Then
    Frame4.Visible = False
    Frame3.Visible = True
    Exit Sub
ElseIf TabStrip1.SelectedItem.Index = 2 Then
Frame4.Visible = True
Frame3.Visible = False
'Frame(fme).Visible = False
'fme = TabStrip1.SelectedItem.Index
End If
End Sub

