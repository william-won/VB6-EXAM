VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "RFID�� �̿��� �ſ���ȸ �ý���"
   ClientHeight    =   10035
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   12015
   LinkTopic       =   "Form1"
   ScaleHeight     =   10035
   ScaleWidth      =   12015
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.Frame Frame4 
      Caption         =   "�ſ���ȸ"
      ForeColor       =   &H00008000&
      Height          =   6135
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   11775
      Begin VB.Frame Frame6 
         Caption         =   "������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   2535
         Left            =   6000
         TabIndex        =   38
         Top             =   3360
         Width           =   5535
         Begin VB.TextBox Text16 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1920
            TabIndex        =   48
            Top             =   1920
            Width           =   3375
         End
         Begin VB.TextBox Text15 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1920
            TabIndex        =   47
            Top             =   1200
            Width           =   3375
         End
         Begin VB.TextBox Text14 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1920
            TabIndex        =   46
            Top             =   480
            Width           =   3375
         End
         Begin VB.Label Label23 
            Caption         =   "��     �� :"
            BeginProperty Font 
               Name            =   "����"
               Size            =   11.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   45
            Top             =   1920
            Width           =   1215
         End
         Begin VB.Label Label22 
            Caption         =   "�����ƿ� :"
            BeginProperty Font 
               Name            =   "����"
               Size            =   11.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   44
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label Label21 
            Caption         =   "����� :"
            BeginProperty Font 
               Name            =   "����"
               Size            =   11.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   43
            Top             =   480
            Width           =   1215
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "�������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   2535
         Left            =   240
         TabIndex        =   35
         Top             =   3360
         Width           =   5415
         Begin VB.TextBox Text13 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1920
            TabIndex        =   42
            Top             =   1920
            Width           =   2895
         End
         Begin VB.TextBox Text12 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1920
            TabIndex        =   41
            Top             =   1200
            Width           =   2895
         End
         Begin VB.TextBox Text11 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1920
            TabIndex        =   40
            Top             =   480
            Width           =   2895
         End
         Begin VB.Label Label20 
            Caption         =   "��Ÿ���� :"
            BeginProperty Font 
               Name            =   "����"
               Size            =   11.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   39
            Top             =   1920
            Width           =   1215
         End
         Begin VB.Label Label19 
            Caption         =   "������� :"
            BeginProperty Font 
               Name            =   "����"
               Size            =   11.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   37
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label Label18 
            Caption         =   "���˰�� :"
            BeginProperty Font 
               Name            =   "����"
               Size            =   11.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   36
            Top             =   480
            Width           =   1215
         End
      End
      Begin VB.TextBox Text10 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8400
         TabIndex        =   34
         Top             =   2760
         Width           =   3135
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         TabIndex        =   32
         Top             =   2760
         Width           =   3375
      End
      Begin VB.TextBox Text9 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8400
         TabIndex        =   30
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4080
         TabIndex        =   28
         Top             =   2040
         Width           =   2655
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         TabIndex        =   26
         Top             =   1320
         Width           =   7935
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8760
         TabIndex        =   25
         Top             =   600
         Width           =   2775
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6360
         TabIndex        =   21
         Top             =   600
         Width           =   1215
      End
      Begin VB.PictureBox Picture1 
         Height          =   2320
         Left            =   240
         Picture         =   "Form1.frx":0000
         ScaleHeight     =   2265
         ScaleWidth      =   1710
         TabIndex        =   18
         Top             =   360
         Width           =   1770
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4080
         TabIndex        =   17
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label17 
         Caption         =   "�ڵ��� :"
         BeginProperty Font 
            Name            =   "����"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7320
         TabIndex        =   33
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "E-Mail :"
         BeginProperty Font 
            Name            =   "����"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   31
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label16 
         Caption         =   "������ :"
         BeginProperty Font 
            Name            =   "����"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7320
         TabIndex        =   29
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label15 
         Caption         =   "�ֹε�Ϲ�ȣ :"
         BeginProperty Font 
            Name            =   "����"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   27
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label14 
         Caption         =   "(����)"
         Height          =   255
         Left            =   8040
         TabIndex        =   24
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label13 
         Caption         =   "(�ѹ�)"
         Height          =   255
         Left            =   5640
         TabIndex        =   23
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label12 
         Caption         =   "(�ѱ�)"
         Height          =   255
         Left            =   3360
         TabIndex        =   22
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "�� �� �� :"
         BeginProperty Font 
            Name            =   "����"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   20
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "��  �� :"
         BeginProperty Font 
            Name            =   "����"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   19
         Top             =   600
         Width           =   855
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   6840
      Top             =   9480
   End
   Begin VB.Frame Frame3 
      Caption         =   "Data Transmit"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   2175
      Left            =   1560
      TabIndex        =   5
      Top             =   7200
      Width           =   5055
      Begin VB.CommandButton Command3 
         Caption         =   "CARD Reading ������ ��ȯ (&A)"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   6
         Top             =   1440
         Width           =   4215
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H000000FF&
         FillStyle       =   0  '�ܻ�
         Height          =   375
         Left            =   4200
         Shape           =   3  '����
         Top             =   480
         Width           =   375
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H0000C000&
         FillStyle       =   0  '�ܻ�
         Height          =   375
         Left            =   2880
         Shape           =   3  '����
         Top             =   480
         Width           =   375
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H0000C0C0&
         FillStyle       =   0  '�ܻ�
         Height          =   375
         Left            =   1680
         Shape           =   3  '����
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label9 
         Alignment       =   2  '��� ����
         Caption         =   "ERROR"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   10
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label8 
         Alignment       =   2  '��� ����
         Caption         =   "RX"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   9
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label7 
         Alignment       =   2  '��� ����
         Caption         =   "TX"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   8
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label5 
         Alignment       =   2  '��� ����
         Caption         =   "���ſ���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   855
      End
      Begin VB.Shape STEP1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  '�ܻ�
         Height          =   375
         Left            =   480
         Shape           =   3  '����
         Top             =   480
         Width           =   375
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "CARD No."
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2175
      Left            =   6840
      TabIndex        =   4
      Top             =   7200
      Width           =   5055
      Begin VB.TextBox Text3 
         Alignment       =   2  '��� ����
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   36
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   4815
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   1680
         Width           =   4815
      End
      Begin VB.Label Label3 
         Caption         =   "Receive Data"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1320
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "RX Bit"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   7200
      Width           =   1335
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�ý�������(&Z)"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   9000
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�ý��ۿ���(&Q)"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   8520
      Width           =   1335
   End
   Begin MSCommLib.MSComm MSComm 
      Left            =   6240
      Top             =   9480
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      RThreshold      =   1
      InputMode       =   1
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "����"
         Size            =   20.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   495
      Left            =   240
      TabIndex        =   14
      Top             =   6480
      Width           =   11535
   End
   Begin VB.Label Label10 
      Alignment       =   1  '������ ����
      Caption         =   "��������������а�  4�г�  �� �� ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6960
      TabIndex        =   11
      Top             =   9600
      Width           =   4935
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  '�ܻ�
      Height          =   255
      Left            =   120
      Top             =   8160
      Width           =   1335
   End
   Begin VB.Menu mnu_file 
      Caption         =   "����"
      Begin VB.Menu mnu_end 
         Caption         =   "����"
      End
   End
   Begin VB.Menu mnu_set 
      Caption         =   "����"
      Begin VB.Menu mnu_port 
         Caption         =   "��Ʈ����"
         Begin VB.Menu mnu_com1 
            Caption         =   "COM1"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnu_com2 
            Caption         =   "COM2"
         End
      End
      Begin VB.Menu mnu_baud 
         Caption         =   "�ӵ�����"
         Begin VB.Menu mnu_9600 
            Caption         =   "9600bps"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnu_115200 
            Caption         =   "115200bps"
         End
      End
   End
   Begin VB.Menu mnu_made 
      Caption         =   "�����ı�"
      Begin VB.Menu mnu_push 
         Caption         =   "�ٿ�! �����ּ���~~"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()    '�ý��� �����ư�� ���� �Լ�
MSComm.PortOpen = True          '��Ʈ�� ����.

Shape1.FillColor = &HFF00&      'Shape1�� ������ �Ķ������� �ٲ۴�(����ON)
Command1.Enabled = False        '�ý��ۿ��� ��ư�� ��Ȱ��ȭ
Command2.Enabled = True         '�ý������� ��ư�� Ȱ��ȭ
Command3.Enabled = True         'ī���б� ��ư�� Ȱ��ȭ
Label2.Caption = "�ý��۰� ������ �Ǿ����ϴ�."      'ȭ�� �Ʒ� �κп� ������ �Ǿ����� ǥ���Ѵ�.

Picture1.Picture = LoadPicture("normal.jpg", 4, 0, 114, 152)    '���� �ʱ�ȭ
Text1.Text = ""                 'Text1 �ʱ�ȭ
Text2.Text = ""                 'Text2 �ʱ�ȭ
Text3.Text = ""                 'Text3 �ʱ�ȭ
Text4.Text = ""                 'Text4 �ʱ�ȭ
Text5.Text = ""                 'Text5 �ʱ�ȭ
Text6.Text = ""                 'Text6 �ʱ�ȭ
Text7.Text = ""                 'Text7 �ʱ�ȭ
Text8.Text = ""                 'Text8 �ʱ�ȭ
Text9.Text = ""                 'Text9 �ʱ�ȭ
Text10.Text = ""                'Text10 �ʱ�ȭ
Text11.Text = ""                'Text11 �ʱ�ȭ
Text12.Text = ""                'Text12 �ʱ�ȭ
Text13.Text = ""                'Text13 �ʱ�ȭ
Text14.Text = ""                'Text14 �ʱ�ȭ
Text15.Text = ""                'Text15 �ʱ�ȭ
Text16.Text = ""                'Text16 �ʱ�ȭ
Label1.Caption = ""             'RX Bit �ʱ�ȭ


End Sub

Private Sub Command2_Click()    '�ý��� ������ư�� ���� �Լ�

MSComm.PortOpen = False         '��Ʈ�� �ݴ´�.
Shape1.FillColor = &HFF&        '���������� �ٲ۴�.(����OFF)
STEP1.FillColor = &HFF&         '���ſ��� LED�ҵ�
Command1.Enabled = True         '�ý��ۿ��� ��ư�� Ȱ��ȭ
Command2.Enabled = False        '�ý������� ��ư�� ��Ȱ��ȭ
Command3.Enabled = False        'ī���б� ��ư�� ��Ȱ��ȭ
Label2.Caption = "�ý��۰� ������ �����Ǿ����ϴ�."      'ȭ�� �Ʒ� �κп� ���� �Ǿ����� ǥ���Ѵ�.

Picture1.Picture = LoadPicture("normal.jpg", 4, 0, 114, 152)    '���� �ʱ�ȭ
Text1.Text = ""                 'Text1 �ʱ�ȭ
Text2.Text = ""                 'Text2 �ʱ�ȭ
Text3.Text = ""                 'Text3 �ʱ�ȭ
Text4.Text = ""                 'Text4 �ʱ�ȭ
Text5.Text = ""                 'Text5 �ʱ�ȭ
Text6.Text = ""                 'Text6 �ʱ�ȭ
Text7.Text = ""                 'Text7 �ʱ�ȭ
Text8.Text = ""                 'Text8 �ʱ�ȭ
Text9.Text = ""                 'Text9 �ʱ�ȭ
Text10.Text = ""                'Text10 �ʱ�ȭ
Text11.Text = ""                'Text11 �ʱ�ȭ
Text12.Text = ""                'Text12 �ʱ�ȭ
Text13.Text = ""                'Text13 �ʱ�ȭ
Text14.Text = ""                'Text14 �ʱ�ȭ
Text15.Text = ""                'Text15 �ʱ�ȭ
Text16.Text = ""                'Text16 �ʱ�ȭ
Label1.Caption = ""             'RX Bit �ʱ�ȭ

End Sub

Private Sub Command3_Click()    '������ ���ۿ� ���� �Լ�
STEP1.FillColor = &HFF00&       '���ſ��� LED����

Picture1.Picture = LoadPicture("normal.jpg", 4, 0, 114, 152)    '���� �ʱ�ȭ
Text1.Text = ""                 'Text1 �ʱ�ȭ
Text2.Text = ""                 'Text2 �ʱ�ȭ
Text3.Text = ""                 'Text3 �ʱ�ȭ
Text4.Text = ""                 'Text4 �ʱ�ȭ
Text5.Text = ""                 'Text5 �ʱ�ȭ
Text6.Text = ""                 'Text6 �ʱ�ȭ
Text7.Text = ""                 'Text7 �ʱ�ȭ
Text8.Text = ""                 'Text8 �ʱ�ȭ
Text9.Text = ""                 'Text9 �ʱ�ȭ
Text10.Text = ""                'Text10 �ʱ�ȭ
Text11.Text = ""                'Text11 �ʱ�ȭ
Text12.Text = ""                'Text12 �ʱ�ȭ
Text13.Text = ""                'Text13 �ʱ�ȭ
Text14.Text = ""                'Text14 �ʱ�ȭ
Text15.Text = ""                'Text15 �ʱ�ȭ
Text16.Text = ""                'Text16 �ʱ�ȭ
Label1.Caption = ""             'RX Bit �ʱ�ȭ

Label2.Caption = "ī�带 �����⿡ ����ּ���."
If ConvertTxt2Binary("2A4963", Form1) = True Then   '�б� �������� ����
End If
End Sub

Private Sub Form_Load()
MSComm.CommPort = 11
MSComm.Settings = "9600" + Setting
Return_Massage = 0
Timer1.Enabled = False
End Sub


Private Sub mnu_push_Click()
    Dim msg, button, title
    msg = "�ȳ��ϼ���~ ^^*" & Chr(10) & Chr(10) & "�� ��ǰ�� ������� ��������� ����־����ϴ�." & Chr(10) & Chr(10) & "���ӿ� �������� ���� ���Դϴ�." & Chr(10) & Chr(10) & "�����մϴ�./(^.^)(_ _)(^.^)/"
    button = vbOKOnly
    title = "�� �� �� ��"
    
    MsgBox msg, button, title
    
End Sub

Private Sub MSComm_OnComm()
    Select Case MSComm.CommEvent
         Case comEvReceive
            Dim buffer() As Byte
            Dim count As Integer
            Dim TempStr As String
            Dim RX_Count As Integer
            TempStr = ""
            
            buffer = MSComm.Input
            
            For count = 0 To UBound(buffer)
                tempchar = Hex(buffer(count))
                If Len(tempchar) = 1 Then tempchar = "0" + tempchar
                TempStr = TempStr + tempchar
            Next count
        
            Text2.Text = Text2.Text + TempStr
            STEP1.FillColor = &HC000&
            RX_Count = RX_Count + 1
            Text3.Text = Mid(Text2.Text, 5, 10)
            Label1.Caption = Len(Text2.Text) / 2
    End Select
                            
            If Len(Text2.Text) / 2 = 8 Then
            
                If Text3.Text = "0F0059373E" Then
                    Picture1.Picture = LoadPicture("oneday798.jpg", 4, 0, 114, 152)
                    Text1.Text = "oneday798"
                    Text4.Text = "������"
                    Text5.Text = "�����"
                    Text6.Text = "Won Dae Hui"
                    Text7.Text = "��⵵ ���ν� ó�α� �ﰡ�� ��Ǫ������ī�� 104-1401ȣ"
                    Text8.Text = "790827-1xxxxxx"
                    Text9.Text = "AB��"
                    Text10.Text = "010-3846-5942"
                    Text11.Text = "������ ����"
                    Text12.Text = "����� �ƴ�"
                    Text13.Text = "����"
                    Text14.Text = "1ȸ(��ӵ���3���ߵ�)"
                    Text15.Text = "1�ƿ�(���ݳ� 0.045%)"
                    Text16.Text = "70��"
                    Label2.Caption = "������ ���� ������ ��ȸ�Ǿ����ϴ�."
                    End If
                
                If Text3.Text = "0F00595CA2" Then
                    Picture1.Picture = LoadPicture("ljh.jpg", 4, 0, 114, 152)
                    Text1.Text = "unc79@hanmail.net"
                    Text4.Text = "������"
                    Text5.Text = "�����"
                    Text6.Text = "Lee Jun Hyuk"
                    Text7.Text = "����� ��ȭ3�� ġ������ ����A 105�� 303ȣ"
                    Text8.Text = "790830-1xxxxxx"
                    Text9.Text = "A��"
                    Text10.Text = "011-9941-6805"
                    Text11.Text = "������ ����"
                    Text12.Text = "����� �ƴ�"
                    Text13.Text = "����"
                    Text14.Text = "����"
                    Text15.Text = "2�ƿ�(���ݳ� 0.04%,0.05%)"
                    Text16.Text = "90��"
                    Label2.Caption = "������ ���� ������ ��ȸ�Ǿ����ϴ�."
                    End If

         End If

 End Sub

