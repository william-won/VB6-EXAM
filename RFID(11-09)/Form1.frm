VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "RFID�� �̿��� �ſ���ȸ �ý���"
   ClientHeight    =   10035
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   12015
   LinkTopic       =   "Form1"
   ScaleHeight     =   10035
   ScaleWidth      =   12015
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.Frame Frame4 
      Caption         =   "�ſ���ȸ"
      ForeColor       =   &H00008000&
      Height          =   6975
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   11775
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
         TabIndex        =   32
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
         TabIndex        =   30
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
         TabIndex        =   28
         Top             =   1320
         Width           =   7815
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
         TabIndex        =   27
         Top             =   600
         Width           =   2655
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
         TabIndex        =   23
         Top             =   600
         Width           =   1215
      End
      Begin VB.PictureBox Picture1 
         Height          =   2320
         Left            =   240
         Picture         =   "Form1.frx":0000
         ScaleHeight     =   2265
         ScaleWidth      =   1710
         TabIndex        =   20
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
         TabIndex        =   19
         Top             =   600
         Width           =   1215
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
         TabIndex        =   31
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
         TabIndex        =   29
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label14 
         Caption         =   "(����)"
         Height          =   255
         Left            =   8040
         TabIndex        =   26
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label13 
         Caption         =   "(�ѹ�)"
         Height          =   255
         Left            =   5640
         TabIndex        =   25
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label12 
         Caption         =   "(�ѱ�)"
         Height          =   255
         Left            =   3360
         TabIndex        =   24
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         Caption         =   "ī���б�(&A)"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3240
         TabIndex        =   7
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '��� ����
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
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   1560
         TabIndex        =   6
         Text            =   "2A4963"
         Top             =   1440
         Width           =   1455
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label6 
         Alignment       =   2  '��� ����
         Caption         =   "ī��˻�"
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
         Left            =   600
         TabIndex        =   9
         Top             =   1560
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
         TabIndex        =   8
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   17
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
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   9600
      Width           =   6255
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
      TabIndex        =   13
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
   Begin VB.Menu mnu_card 
      Caption         =   "ī��"
      Begin VB.Menu mnu_creat 
         Caption         =   "���"
      End
      Begin VB.Menu mnu_del 
         Caption         =   "����"
      End
      Begin VB.Menu mnu_reg 
         Caption         =   "��ϵ�ī��"
      End
   End
   Begin VB.Menu mnu_help 
      Caption         =   "����"
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
End Sub

Private Sub Command2_Click()    '�ý��� ������ư�� ���� �Լ�
MSComm.PortOpen = False         '��Ʈ�� �ݴ´�.
Shape1.FillColor = &HFF&        '���������� �ٲ۴�.(����OFF)
Command1.Enabled = True         '�ý��ۿ��� ��ư�� Ȱ��ȭ
Command2.Enabled = False        '�ý������� ��ư�� ��Ȱ��ȭ
Command3.Enabled = False        'ī���б� ��ư�� ��Ȱ��ȭ
Label2.Caption = "�ý��۰� ������ �����Ǿ����ϴ�."      'ȭ�� �Ʒ� �κп� ���� �Ǿ����� ǥ���Ѵ�.
End Sub

Private Sub Command3_Click()    '������ ���ۿ� ���� �Լ�
STEP1.FillColor = &HFF00&       '���ſ��� LED����
Picture1.Picture = LoadPicture("normal.jpg", 4, 0, 114, 152)    '���� �ʱ�ȭ
Text2.Text = ""                 'Text2 �ʱ�ȭ
Text3.Text = ""                 'Text3 �ʱ�ȭ
Text4.Text = ""                 'Text4 �ʱ�ȭ
Text5.Text = ""                 'Text5 �ʱ�ȭ
Text6.Text = ""                 'Text6 �ʱ�ȭ
Text7.Text = ""                 'Text7 �ʱ�ȭ
Text8.Text = ""                 'Text8 �ʱ�ȭ
Text9.Text = ""                 'Text9 �ʱ�ȭ
Label1.Caption = ""             'RX Bit �ʱ�ȭ
Label2.Caption = "ī�带 �����⿡ ����ּ���."
If ConvertTxt2Binary(Text1.Text, Form1) = True Then   'READ
End If
End Sub

Private Sub Command5_Click()
Timer1.Enabled = True
End Sub

Private Sub Form_Load()
MSComm.CommPort = 1
MSComm.Settings = "9600" + Setting
Return_Massage = 0
Timer1.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub


Private Sub Image1_Click()

End Sub

Private Sub mnu_creat_Click()
Reg_Card = InputBox("����� ī���ȣ�� �Է��Ͻʽÿ�")
End Sub

Private Sub mnu_del_Click()
a = MsgBox("��ϵ� ī�带 ���� �Ͻðڽ��ϱ�?", vbYesNo)
If a = 6 Then
   Reg_Card = ""
End If
End Sub

Private Sub mnu_end_Click()
End
End Sub


Private Sub mnu_reg_Click()
MsgBox (Reg_Card)
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
                    Text4.Text = "������"
                    Text5.Text = "�����"
                    Text6.Text = "Won Dae Hee"
                    Text7.Text = "��⵵ �Ⱦ�� ���ȱ� �Ⱦ�5�� 710-87���� ������ B-302ȣ"
                    Text8.Text = "790827-1xxxxxx"
                    Text9.Text = "AB"
                    Label2.Caption = "������ ���� ������ ��ȸ�Ǿ����ϴ�."
                Else
                    Label2.Caption = "�ſ���ȸ������ ���ų� ī���б� �����Դϴ�. �ٽ� �õ����ּ���."
                End If
            End If
End Sub

