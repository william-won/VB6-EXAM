VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00800000&
   BorderStyle     =   0  '없음
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":000C
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtDataBuf 
      Height          =   285
      Left            =   13710
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   1920
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Frame frameOBIS 
      BackColor       =   &H00200005&
      Caption         =   "OBIS Code Parameters"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1395
      Left            =   120
      TabIndex        =   8
      Top             =   360
      Width           =   15105
      Begin VB.TextBox txtOBIS 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   1380
         MaxLength       =   3
         TabIndex        =   25
         Text            =   "0"
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txtOBIS 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   1890
         MaxLength       =   3
         TabIndex        =   24
         Text            =   "0"
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txtOBIS 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   2400
         MaxLength       =   3
         TabIndex        =   23
         Text            =   "1"
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txtOBIS 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   2910
         MaxLength       =   3
         TabIndex        =   22
         Text            =   "0"
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txtOBIS 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   4
         Left            =   3420
         MaxLength       =   3
         TabIndex        =   21
         Text            =   "0"
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txtOBIS 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   5
         Left            =   3930
         MaxLength       =   3
         TabIndex        =   20
         Text            =   "255"
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txtClassID 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5460
         MaxLength       =   2
         TabIndex        =   19
         Text            =   "8"
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txtAttrID 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7290
         MaxLength       =   2
         TabIndex        =   18
         Text            =   "2"
         Top             =   600
         Width           =   495
      End
      Begin VB.CommandButton cmdRead_XXX 
         Caption         =   "&Read XXX"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   11550
         TabIndex        =   17
         Top             =   360
         Width           =   990
      End
      Begin VB.CommandButton cmdWrite_XXX 
         Caption         =   "&Write XXX"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   11550
         TabIndex        =   16
         Top             =   810
         Width           =   990
      End
      Begin VB.CommandButton cmdQuit 
         Caption         =   "&Quit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   13890
         TabIndex        =   15
         Top             =   810
         Width           =   990
      End
      Begin VB.TextBox txtVZ 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   330
         Left            =   8280
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   14
         Text            =   "0"
         Top             =   600
         Width           =   495
      End
      Begin VB.CommandButton cmdReadVZ 
         Caption         =   "&VZ Read"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   8880
         TabIndex        =   13
         Top             =   480
         Width           =   645
      End
      Begin VB.CommandButton cmdCOMset 
         Caption         =   "COM &Set"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   12720
         TabIndex        =   12
         Top             =   360
         Width           =   990
      End
      Begin VB.ComboBox cmbAssoc 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   10500
         TabIndex        =   11
         Text            =   "Combo1"
         Top             =   600
         Width           =   825
      End
      Begin VB.CommandButton cmdPassword 
         Caption         =   "&PW Set"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   12720
         TabIndex        =   10
         Top             =   810
         Width           =   990
      End
      Begin VB.CommandButton cmdSelectiveLP 
         Caption         =   "SelectLP"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   13890
         TabIndex        =   9
         Top             =   360
         Width           =   990
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  '투명
         Caption         =   "OBIS Code :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   150
         TabIndex        =   30
         Top             =   630
         Width           =   1275
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  '투명
         Caption         =   "Class ID :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   4500
         TabIndex        =   29
         Top             =   630
         Width           =   1005
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  '투명
         Caption         =   "Attribute ID :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   6030
         TabIndex        =   28
         Top             =   630
         Width           =   1305
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  '투명
         Caption         =   "VZ :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   5
         Left            =   7860
         TabIndex        =   27
         Top             =   630
         Width           =   435
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  '투명
         Caption         =   "Assoc. :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   6
         Left            =   9690
         TabIndex        =   26
         Top             =   630
         Width           =   1005
      End
   End
   Begin VB.Frame fraSelect 
      BackColor       =   &H00200005&
      Caption         =   "OBIS List Selection"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   7890
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   5745
      Begin VB.ListBox lstOBIS 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7200
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   5445
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00200005&
      Caption         =   "Data Response"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   7890
      Left            =   6000
      TabIndex        =   0
      Top             =   1920
      Width           =   9255
      Begin RichTextLib.RichTextBox rtxInfoBox 
         Height          =   6885
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   8955
         _ExtentX        =   15796
         _ExtentY        =   12144
         _Version        =   393217
         BorderStyle     =   0
         ScrollBars      =   2
         TextRTF         =   $"frmMain.frx":5361
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  '투명
         Caption         =   "Object Name :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   5
         Top             =   330
         Width           =   1395
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  '투명
         Caption         =   "Attribute :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   4410
         TabIndex        =   4
         Top             =   330
         Width           =   1065
      End
      Begin VB.Label lblObjOut 
         BackStyle       =   0  '투명
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1530
         TabIndex        =   3
         Top             =   330
         Width           =   2805
      End
      Begin VB.Label lblAttrOut 
         BackStyle       =   0  '투명
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5490
         TabIndex        =   2
         Top             =   330
         Width           =   3645
      End
   End
   Begin VB.Image imgPowerOff 
      Height          =   1950
      Left            =   13560
      Picture         =   "frmMain.frx":53F7
      Top             =   9600
      Width           =   1800
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*                                                                                                  *

Private Sub Form_Load()
    
    Dim ii_Index As Byte
    
    Call Init_DataType_CONST_Value
    
    Call Init_Comm_Setting_Value
    
    Call Init_OBIS_Table
    
    Call Disp_OBIS_Table
    
    cmbAssoc.Clear
    For ii_Index = 0 To 5 Step 1
        cmbAssoc.AddItem CStr(ii_Index)
    Next ii_Index
    cmbAssoc.ListIndex = 0
    
End Sub
'*                                                                                                  *

'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    If MsgBox("Do you want to quit DLMS Read/Write Tester now?", vbYesNo, "Confirm Quit") = vbNo Then
'        Cancel = True
'        Exit Sub
'    End If
'End Sub
'*                                                                                                  *

Private Sub Form_Unload(Cancel As Integer)
    Set frmMain = Nothing
    End
End Sub
'*                                                                                                  *

Private Sub cmdQuit_Click()
    Unload Me
End Sub
'*                                                                                                  *

Private Sub cmdCOMset_Click()
    
    Dim tSTR As String
    Dim StrDIR As String
    Dim FileNum
    
    tSTR = InputBox("Input COM port number for communication.", "COM Port Input", sCommSET.COM_Port)
    If tSTR = "" Then
        Exit Sub
    End If
    If IsNumeric(tSTR) = False Then
        MsgBox "Input COM port number was not numeric !", vbExclamation, "Wrong Input"
        Exit Sub
    End If
    If Val(tSTR) < 1 Or Val(tSTR) > 255 Then
        MsgBox "Input COM port number was over range(COM 1~255) !", vbExclamation, "Wrong Input"
        Exit Sub
    End If
    
    sCommSET.COM_Port = CByte(tSTR)
    
    tSTR = InputBox("Input device number for communication." & _
                    vbNewLine & "0-Direct Cable, 1-Optic", "COMM Device Input", sCommSET.Device)
    If tSTR = "" Then
        Exit Sub
    End If
    If IsNumeric(tSTR) = False Then
        MsgBox "Input device number was not numeric !", vbExclamation, "Wrong Input"
        Exit Sub
    End If
    If Val(tSTR) < 0 Or Val(tSTR) > 1 Then
        MsgBox "Input device number was over range(0~1) !", vbExclamation, "Wrong Input"
        Exit Sub
    End If
        
    sCommSET.Device = CByte(tSTR)
    
    With sCommSET
        If .Device = 1 Then     'Optic
            .Baud_Rate = 300
            .Parity_Bit = 2 'EvenParity
        Else
            .Baud_Rate = 9600   'Direct Cable
            .Parity_Bit = 0 'NoneParity
        End If
    End With
        
    StrDIR = App.Path & "\CommSet.dat"
    With sCommSET
        FileNum = FreeFile
        Open StrDIR For Output As #FileNum
            Print #FileNum, CStr(.COM_Port)
            Print #FileNum, CStr(.Baud_Rate)
            Print #FileNum, CStr(.Parity_Bit)
            Print #FileNum, CStr(.Device)
        Close #FileNum
    End With
    
End Sub
'*                                                                                                  *

Private Sub cmdPassword_Click()
    
    frmPassword.Show 1
    
    Call cmbAssoc_Click     '현재 combo index에 맞는 password를 바로 적용
    
End Sub
'*                                                                                                  *

Private Sub cmbAssoc_Click()

    If cmbAssoc.ListIndex = 0 Then  '''''''''Association 1''''''''''
        gClientID = &H10
        gContext = 1
        gConformance = &H1819
        gAuthenication_Mech = 0
        assoc_index.lls_secret = "        "
    ElseIf cmbAssoc.ListIndex = 1 Then '''''''''Association 2''''''''''
        gClientID = &H11
        gContext = 1
        gConformance = &H1819
        gAuthenication_Mech = 1
        assoc_index.lls_secret = gPASS_Manufacture
    ElseIf cmbAssoc.ListIndex = 2 Then '''''''''Association 3''''''''''
        MsgBox "This association is for factory access mode only !", vbExclamation, "Association not allowed"
        gClientID = &H12
        gContext = 1
        gConformance = &H1819
        gAuthenication_Mech = 1
        assoc_index.lls_secret = gPASS_Factory
    ElseIf cmbAssoc.ListIndex = 3 Then
        gClientID = &H13
        gContext = 2
        gConformance = &H180000
        gAuthenication_Mech = 0
        assoc_index.lls_secret = "        "
    ElseIf cmbAssoc.ListIndex = 4 Then
        gClientID = &H14
        gContext = 2
        gConformance = &H180000
        gAuthenication_Mech = 1
        assoc_index.lls_secret = gPASS_Manufacture
    ElseIf cmbAssoc.ListIndex = 5 Then
        MsgBox "This association is for factory access mode only !", vbExclamation, "Association not allowed"
        gClientID = &H15
        gContext = 2
        gConformance = &H180000
        gAuthenication_Mech = 1
        assoc_index.lls_secret = gPASS_Factory
    End If

End Sub
'*                                                                                                  *

Private Sub cmdWrite_XXX_Click()
        
    Dim myOBIS As OBISCODE
    Dim data_length_set1 As Long
    Dim buffer1() As Byte
    Dim data_index_set1 As Byte
    Dim data_type_set1 As Byte
    Dim attr_index1 As Byte
    Dim class_id1 As Byte
    Dim tVal As Long
    Dim ii_Index As Integer
    
    Dim tBYTE As Byte
    Dim tSTR As String
    
    If Check_Support_Set = False Then
        MsgBox "This OBIS code cannot be set !", vbExclamation, "Get only OBIS"
        cmdWrite_XXX.Enabled = False
        Exit Sub
    End If
    
    Call Enable_ControlButton(False)
    DoEvents
    
    'MsgBox "Data writing function is not supported now...", vbInformation, "Not Support"

    rtxInfoBox.Text = ""
    With myOBIS
        .a = CByte(txtOBIS(0).Text):    .b = CByte(txtOBIS(1).Text)
        .c = CByte(txtOBIS(2).Text):    .d = CByte(txtOBIS(3).Text)
        .e = CByte(txtOBIS(4).Text):    .f = CByte(txtOBIS(5).Text)
    End With
    class_id1 = CByte(txtClassID.Text): attr_index1 = CByte(txtAttrID.Text)
    
    ii_Index = lstOBIS.ListIndex
    
    data_type_set1 = CLng(sOBIS_Tbl(ii_Index).SetType)
    data_length_set1 = CLng(sOBIS_Tbl(ii_Index).SetLen)
    data_index_set1 = 255   'data index set not used as data type is not array

    gSet_DataType = data_type_set1
    gSet_DataLen = data_length_set1
    
    If (myOBIS.a = 1 And myOBIS.b = 0 And myOBIS.c = 0 And _
            myOBIS.d = 239 And myOBIS.e = 0 And myOBIS.f = 255) Then
        frmMemorySet.Show 1
    Else
        frmDataSet.Show 1
    End If
    
    If Set_Confirmed_Or_Not = False Then GoTo DO_FINAL_JOB
    
'    gClientID = &H11
'    gContext = 1
'    gConformance = &H1819
'    gAuthenication_Mech = 1
    
    If PortOpen = False Then GoTo DO_FINAL_JOB
    
    If SNRMSend = 0 Then
        If Assoc = 0 Then
            tVal = write_XXX(myOBIS, attr_index1, class_id1, ls_buf(0), data_length_set1, _
                            data_type_set1, data_index_set1)
        Else
            tVal = -1
        End If
    Else
        tVal = -1
    End If
    
    Call disconnect
    
    If tVal = 0 Then
        rtxInfoBox.Text = rtxInfoBox.Text + vbCrLf + "Data Written !"
    Else
        rtxInfoBox.Text = rtxInfoBox.Text + vbCrLf + "Data Write Failed..."
    End If

DO_FINAL_JOB:

    Call Enable_ControlButton(True)
    
End Sub
'*                                                                                                  *

Private Sub cmdReadVZ_Click()

    Dim DATA_DISPLAY As Byte
    Dim tOutStr As String
    Dim xy As Long
    Dim myOBIS As OBISCODE
    Dim attr_index1 As Byte
    Dim class_id1 As Byte
    Dim Read_Data_Len As Long
    Dim nVal As Long
    Dim Num_Struct As Long
    Dim Arr_Count As Byte
    Dim abc As Integer
    Dim def As Long
    Dim ij As Integer

    Call Enable_ControlButton(False)
    Call Clear_LabelData

    With myOBIS
        .a = 1: .b = 0: .c = 0
        .d = 1: .e = 0: .f = 255
    End With
    class_id1 = 1
    attr_index1 = 2
    
    
    Erase ls_buf
    Read_Data_Len = 0
    DATA_DISPLAY = 0
    Num_Struct = 0
    
    If PortOpen = False Then GoTo DO_FINAL_JOB
    
    If SNRMSend = 0 Then
        If Assoc = 0 Then
            nVal = read_XXX(myOBIS, attr_index1, class_id1, Read_Data_Len, sAccess, 0)
        Else
            Read_Data_Len = 0
        End If
    Else
        Read_Data_Len = 0
    End If
    
    Call disconnect
    
    If Read_Data_Len = 0 Then
        tOutStr = "GET FAILURE"
        rtxInfoBox.Text = rtxInfoBox.Text + vbCrLf + tOutStr
        Call Enable_ControlButton(True)
        Exit Sub
    Else
        gVZ = ls_buf(0)
        txtVZ.Text = CByte(gVZ)
    End If
    
    tOutStr = "VZ = " & HexToTwo(Hex(ls_buf(0)))
    rtxInfoBox.Text = rtxInfoBox.Text + vbCrLf + tOutStr
    
    Already_Read_VZ_Or_Not = True

DO_FINAL_JOB:
    
    Call Enable_ControlButton(True)

End Sub
'*                                                                                                  *

Private Sub cmdSelectiveLP_Click()

    Dim DATA_DISPLAY As Byte
    Dim tOutStr As String
    Dim xy As Long
    Dim myOBIS As OBISCODE
    Dim attr_index1 As Byte
    Dim class_id1 As Byte
    Dim Read_Data_Len As Long
    Dim nVal As Long
    Dim Num_Struct As Long
    Dim Arr_Count As Byte
    Dim abc As Integer
    Dim def As Long
    Dim ij As Integer
    
    Call Enable_ControlButton(False)
    Call Clear_LabelData

    With myOBIS
        .a = 1:     .b = 128:   .c = 128
        .d = 128:   .e = 11:    .f = 255
    End With
    class_id1 = 1
    attr_index1 = 2
        
    Erase ls_buf
    Read_Data_Len = 0
    DATA_DISPLAY = 0
    Num_Struct = 0
    
    If PortOpen = False Then GoTo DO_FINAL_JOB
    
    If SNRMSend = 0 Then
        If Assoc = 0 Then
            nVal = read_XXX(myOBIS, attr_index1, class_id1, Read_Data_Len, sAccess, 0)
        Else
            Read_Data_Len = 0
        End If
    Else
        Read_Data_Len = 0
    End If
    
    Call disconnect
    
    If Read_Data_Len = 0 Then
        tOutStr = "LP Index GET FAILURE"
        rtxInfoBox.Text = rtxInfoBox.Text + vbCrLf + tOutStr
        Call Enable_ControlButton(True)
        Exit Sub
    Else
        Max_LP_Index = CLng(ls_buf(0)) * 256 + CLng(ls_buf(1))
    End If
    
    Call Enable_ControlButton(True)
    
    tOutStr = HexToTwo(Hex(ls_buf(0))) & " " & HexToTwo(Hex(ls_buf(1))) & _
                " : Max LP Index = " & CStr(Max_LP_Index)
    rtxInfoBox.Text = rtxInfoBox.Text + vbCrLf + tOutStr
    
    If Max_LP_Index = 0 Then
        MsgBox "There is no LP in the meter !", vbInformation, "No LP"
        GoTo DO_FINAL_JOB
    End If
    
    Selective_Or_Not = False
    
    frmSelective.Show 1
    
    If Selective_Or_Not = False Then GoTo DO_FINAL_JOB
    
    txtOBIS(0).Text = 1:    txtOBIS(1).Text = 0:    txtOBIS(2).Text = 99
    txtOBIS(3).Text = 1:    txtOBIS(4).Text = 0:    txtOBIS(5).Text = 255
    txtClassID.Text = 7:    txtAttrID.Text = 2
    
    Call cmdRead_XXX_Click
    
    Selective_Or_Not = False
    
    Exit Sub
    
DO_FINAL_JOB:
    
    Call Enable_ControlButton(True)
    
End Sub
'*                                                                                                  *

Private Sub cmdRead_XXX_Click()

    Dim DATA_DISPLAY As Byte
    Dim tOutStr As String
    Dim xy As Long
    Dim obis As OBISCODE
    Dim attr_index1 As Byte
    Dim class_id1 As Byte
    Dim Read_Data_Len As Long
    Dim nVal As Long
    Dim Num_Struct As Long
    Dim Arr_Count As Integer
    Dim abc As Integer
    Dim def As Long
    Dim ij As Integer
    
    Dim tCALC(11) As Byte    'Unsigned pulse 환산을 위해
    Dim tSNG(3) As Byte
    Dim tFLOAT As Single
    Dim tTEMP As Long
    
    Dim read_data_length1 As Long    'test
    Dim num_elements_array As Integer

    If Now_Doing_Comm = True Then Exit Sub
    
    If (txtOBIS(0).Text = "-") Or (txtOBIS(1).Text = "-") Or (txtOBIS(2).Text = "-") Or _
        (txtOBIS(3).Text = "-") Or (txtOBIS(4).Text = "-") Or (txtOBIS(5).Text = "-") Then
            Exit Sub
    End If
    
    Call Enable_ControlButton(False)
    Call Clear_LabelData

'    gClientID = &H10
'    gContext = 1
'    gConformance = &H1819
'    gAuthenication_Mech = 0

    obis.a = CByte(txtOBIS(0).Text)
    obis.b = CByte(txtOBIS(1).Text)
    obis.c = CByte(txtOBIS(2).Text)
    obis.d = CByte(txtOBIS(3).Text)
    obis.e = CByte(txtOBIS(4).Text)
    obis.f = CByte(txtOBIS(5).Text)
    attr_index1 = CByte(txtAttrID.Text)
    class_id1 = CByte(txtClassID.Text)
    
    Erase ls_buf
    Read_Data_Len = 0:      tOutStr = ""
    DATA_DISPLAY = 0:       Num_Struct = 0
    
    If PortOpen = False Then GoTo DO_FINAL_JOB
    
    If SNRMSend = 0 Then
        If Assoc = 0 Then
            nVal = read_XXX(obis, attr_index1, class_id1, Read_Data_Len, _
                            sAccess, IIf(Selective_Or_Not = True, CByte(1), CByte(0)))
            Call disconnect
        Else
            Read_Data_Len = 0
        End If
    Else
        Read_Data_Len = 0
    End If
    'Call ClosePort
    
    If Read_Data_Len = 0 Then
        tOutStr = "GET FAILURE"
        rtxInfoBox.Text = rtxInfoBox.Text + vbCrLf + tOutStr
        Call Enable_ControlButton(True)
        Call disconnect
        Exit Sub
    End If
    
    DATA_DISPLAY = 0
    DT_ARRAY = 1
    xy = 0
            
    Select Case (class_id1)
        Case 1
            lblObjOut.Caption = "DATA"
            If attr_index1 = 2 Then lblAttrOut.Caption = "Value"
    
        Case 3
            lblObjOut.Caption = "REGISTER"
            If attr_index1 = 3 Then
                lblAttrOut.Caption = "Scaler Unit"
                tOutStr = "Scaler is " & HexToTwo(CStr(Hex(ls_buf(0)))) & vbCrLf
                tOutStr = tOutStr + "Unit is " & HexToTwo(CStr(Hex(ls_buf(1)))) & vbCrLf
                DATA_DISPLAY = 1
            ElseIf attr_index1 = 2 Then lblAttrOut.Caption = "Value"
            End If
        Case 4
            lblObjOut.Caption = "EXTENDED REGISTER"
            If attr_index1 = 3 Then
                lblAttrOut.Caption = "Scaler Unit"
                tOutStr = "Scaler is " & HexToTwo(CStr(Hex(ls_buf(0)))) & vbCrLf
                tOutStr = tOutStr + "Unit is " & HexToTwo(CStr(Hex(ls_buf(1)))) & vbCrLf
                DATA_DISPLAY = 1
            ElseIf attr_index1 = 2 Then lblAttrOut.Caption = "Value"
            ElseIf attr_index1 = 4 Then lblAttrOut.Caption = "Status"
            ElseIf attr_index1 = 5 Then lblAttrOut.Caption = "Capture time"
            End If
        Case 5
            lblObjOut.Caption = "DEMAND REGISTER"
            If attr_index1 = 4 Then
                lblAttrOut.Caption = "Scaler Unit"
                tOutStr = "Scaler is " & HexToTwo(CStr(Hex(ls_buf(0)))) & vbCrLf
                tOutStr = tOutStr + "Unit is " & HexToTwo(CStr(Hex(ls_buf(1)))) & vbCrLf
                DATA_DISPLAY = 1
            ElseIf attr_index1 = 2 Then lblAttrOut.Caption = "Current Average Value"
            ElseIf attr_index1 = 3 Then lblAttrOut.Caption = "Last Average Value"
            ElseIf attr_index1 = 5 Then lblAttrOut.Caption = "Status"
            ElseIf attr_index1 = 6 Then lblAttrOut.Caption = "Capture time"
            ElseIf attr_index1 = 7 Then lblAttrOut.Caption = "Start time Current"
            ElseIf attr_index1 = 8 Then lblAttrOut.Caption = "Period"
            ElseIf attr_index1 = 9 Then lblAttrOut.Caption = "Number of periods"
            End If
    
        Case 7 ' profile
            lblObjOut.Caption = "PROFILE"
            If attr_index1 = 2 Then
                xy = 0
                lblAttrOut.Caption = "Buffer Values"
                'tOutStr = "Buffer Values are : "
                If ls_buf(xy) = DT_ARRAY Then
                    xy = xy + 1
                    If ls_buf(xy) = &H82 Then
                        xy = xy + 1
                        Arr_Count = (ls_buf(xy) * &H100&) Or (ls_buf(xy + 1) And &HFF&)
                        xy = xy + 2
                    Else
                        Arr_Count = ls_buf(xy)
                        xy = xy + 1
                    End If
                    'Arr_Count = ls_buf(xy)
                    'xy = xy + 1
                    For abc = 0 To Arr_Count - 1
                        tOutStr = tOutStr + vbCrLf + vbCrLf + "Profile Entry  " + HexToTwo(CStr(Hex(abc))) + vbCrLf
                        If ls_buf(xy) = DT_STRUCTURE Then
                            xy = xy + 1
                            Num_Struct = ls_buf(xy)
                            xy = xy + 1
                        End If
                        For ij = 0 To Num_Struct - 1
                            data_type1 = ls_buf(xy)
                            xy = xy + 1
                            If find_data_length1() = 0 Then
                                data_length1 = ls_buf(xy)
                                xy = xy + 1
                            End If
                            'tOutStr = tOutStr + vbCrLf + "Value of Capture Object " + HexToTwo(CStr(Hex(ij))) + " is "
                            tOutStr = tOutStr + vbCrLf + "CapObj " + HexToTwo(CStr(Hex(ij))) + " is "
                            For def = 0 To data_length1 - 1
                                tOutStr = tOutStr + HexToTwo(CStr(Hex(ls_buf(xy)))) + " "
                                tCALC(def) = ls_buf(xy)     'Unsigned Pulse환산을 위해...
                                xy = xy + 1
                            Next def
                            
                            Select Case data_length1        'Unsigned Pulse환산을 위해...
                                Case 4
                                    If tCALC(0) > &HF Then     'Float가 아님
                                        For def = 0 To 3 Step 1
                                            tSNG(3 - def) = tCALC(def)
                                        Next def
                                        CopyMemory tFLOAT, tSNG(0), 4
                                        tOutStr = tOutStr & " = " & Format(tFLOAT, "#0.0000")
                                    Else
                                        tOutStr = tOutStr & " = " & CStr(Val(tCALC(0)) * CDbl(256 ^ 3) + _
                                            Val(tCALC(1)) * CDbl(256 ^ 2) + _
                                            Val(tCALC(2)) * CDbl(256 ^ 1) + _
                                            Val(tCALC(3)))
                                    End If
                                Case 2
                                    tOutStr = tOutStr & " = " & CStr(CLng(tCALC(0)) * 256 + CLng(tCALC(1)))
                                Case 1
                                    tOutStr = tOutStr & " = " & CStr(tCALC(0))
                                Case 12
                                    tTEMP = CLng(tCALC(0)) * 256 + CLng(tCALC(1))
                                    If tTEMP > 9999 Then
                                        tOutStr = tOutStr & " = " & "xxxx-"
                                    Else
                                        tOutStr = tOutStr & " = " & CStr(tTEMP) & "-"
                                    End If
                                    If tCALC(2) > 12 Then
                                        tOutStr = tOutStr & "xx-"
                                    Else
                                        tOutStr = tOutStr & Format(tCALC(2), "00") & "-"
                                    End If
                                    If tCALC(3) > 31 Then
                                        tOutStr = tOutStr & "xx "
                                    Else
                                        tOutStr = tOutStr & Format(tCALC(3), "00") & " "
                                    End If
                                    If tCALC(5) > 23 Then
                                        tOutStr = tOutStr & "xx:"
                                    Else
                                        tOutStr = tOutStr & Format(tCALC(5), "00") & ":"
                                    End If
                                    If tCALC(6) > 59 Then
                                        tOutStr = tOutStr & "xx:"
                                    Else
                                        tOutStr = tOutStr & Format(tCALC(6), "00") & ":"
                                    End If
                                    If tCALC(7) > 59 Then
                                        tOutStr = tOutStr & "xx"
                                    Else
                                        tOutStr = tOutStr & Format(tCALC(7), "00")
                                    End If
                            End Select      'Unsigned 환산 여기까지
                        
                        Next ij
                    Next abc
                    'tOutStr = tOutStr + vbCrLf
                    DATA_DISPLAY = 1
                End If
                
            ElseIf attr_index1 = 3 Then
                lblAttrOut.Caption = " Capture objects "
                tOutStr = "Number of capture objects is  " & HexToTwo(CStr(Hex(Read_Data_Len / 11))) & vbCrLf
                xy = 0
                For ij = 0 To (Read_Data_Len / 11) - 1
                    tOutStr = tOutStr + vbCrLf & vbCrLf + "     Capture Object " & HexToTwo(CStr(Hex(ij))) & vbCrLf & vbCrLf
                    tOutStr = tOutStr + "Class ID :" & HexToTwo(CStr(Hex(ls_buf(xy)))) & " " & HexToTwo(CStr(Hex(ls_buf(xy + 1)))) & vbCrLf
                    tOutStr = tOutStr + vbCrLf + "OBIS code : " & HexToTwo(CStr(Hex(ls_buf(xy + 2)))) & " " & HexToTwo(CStr(Hex(ls_buf(xy + 3)))) & " " & HexToTwo(CStr(Hex(ls_buf(xy + 4)))) & " " & HexToTwo(CStr(Hex(ls_buf(xy + 5)))) & " " & HexToTwo(CStr(Hex(ls_buf(xy + 6)))) & " " & HexToTwo(CStr(Hex(ls_buf(xy + 7)))) & vbCrLf
                    tOutStr = tOutStr + "Attribute Index " & HexToTwo(CStr(Hex(ls_buf(xy + 8)))) + vbCrLf
                    tOutStr = tOutStr + "Data Index :" & HexToTwo(CStr(Hex(ls_buf(xy + 9)))) & " " & HexToTwo(CStr(Hex(ls_buf(xy + 10))))
                    xy = xy + 11
                Next ij
                DATA_DISPLAY = 1
            ElseIf attr_index1 = 4 Then lblAttrOut.Caption = "Capture period"
            ElseIf attr_index1 = 5 Then lblAttrOut.Caption = "Sort method"
            ElseIf attr_index1 = 6 Then lblAttrOut.Caption = "Sort object"
            ElseIf attr_index1 = 7 Then lblAttrOut.Caption = "Entries in use"
            ElseIf attr_index1 = 8 Then lblAttrOut.Caption = "Profile entries"
            End If
        Case 8  '  clock
                lblObjOut.Caption = "CLOCK"
                If attr_index1 = 2 Then
                    lblAttrOut.Caption = "Time"
                ElseIf attr_index1 = 3 Then lblAttrOut.Caption = "Time zone"
                ElseIf attr_index1 = 4 Then lblAttrOut.Caption = "Status"
                ElseIf attr_index1 = 5 Then lblAttrOut.Caption = "Day Light savings begin"
                ElseIf attr_index1 = 6 Then lblAttrOut.Caption = "Day Light savings end"
                ElseIf attr_index1 = 7 Then lblAttrOut.Caption = "Day Light savings deviation"
                ElseIf attr_index1 = 8 Then lblAttrOut.Caption = "Day Light savings enabled"
                ElseIf attr_index1 = 9 Then lblAttrOut.Caption = "Clock base"
                End If
        Case 9
                lblObjOut.Caption = "SCRIPT TABLE"
                If attr_index1 = 2 Then
                    xy = 0
                    lblAttrOut.Caption = "Script"
                    tOutStr = "Array of " & HexToTwo(CStr(Hex(Read_Data_Len / 2))) & " scripts" & vbCrLf
                    For ij = 0 To (Read_Data_Len / 2) - 1
                        tOutStr = tOutStr + "Script Identifier " & HexToTwo(CStr(Hex(ij))) & " is "
                        tOutStr = tOutStr + HexToTwo(CStr(Hex(ls_buf(xy))))
                        xy = xy + 1
                        tOutStr = tOutStr + HexToTwo(CStr(Hex(ls_buf(xy)))) & vbCrLf
                        xy = xy + 1
                    Next ij
                    DATA_DISPLAY = 1
                End If
        Case 11
            lblObjOut.Caption = "SPECIAL DAYS TABLE"
            If attr_index1 = 2 Then
                xy = 0
                lblAttrOut.Caption = "Special day Entry"
                tOutStr = "Array of " & HexToTwo(CStr(Hex(Read_Data_Len / 8))) & " special day entries" & vbCrLf
                For ij = 0 To (Read_Data_Len / 8) - 1
                    tOutStr = tOutStr + vbCrLf + "-------Special Day Entry  " & HexToTwo(CStr(Hex(ij))) & vbCrLf
                    tOutStr = tOutStr + "Index is " & HexToTwo(CStr(Hex(ls_buf(xy))))
                    xy = xy + 1
                    tOutStr = tOutStr + " " & HexToTwo(CStr(Hex(ls_buf(xy))))
                    xy = xy + 1
                    tOutStr = tOutStr + vbCrLf + "Special Day date is "
                    For abc = 0 To 4
                        tOutStr = tOutStr + HexToTwo(CStr(Hex(ls_buf(xy)))) & " "
                        xy = xy + 1
                    Next abc
                    tOutStr = tOutStr + vbCrLf + "Day ID is " & HexToTwo(CStr(Hex(ls_buf(xy))))
                    xy = xy + 1
                Next ij
                DATA_DISPLAY = 1
            End If
        Case 12
            lblObjOut.Caption = "Association SN"

        Case 15
'Below old codes are Sep.27 version
'            xy = 0
'            lblObjOut.Caption = "Association LN"
'            If attr_index1 = 3 Then
'                tOutStr = "Client SAP is" & HexToTwo(CStr(Hex(ls_buf(xy)))) & vbCrLf
'                tOutStr = tOutStr + "Server SAP is " & HexToTwo(CStr(Hex(ls_buf(xy + 1))))
'                DATA_DISPLAY = 1
'            ElseIf attr_index1 = 4 Then lblAttrOut.Caption = "Application Context Name"
'            ElseIf attr_index1 = 5 Then
'                tOutStr = "Conformance Block is " & HexToTwo(CStr(Hex(ls_buf(xy)))) & " " & HexToTwo(CStr(Hex(ls_buf(xy + 1)))) & " " & HexToTwo(CStr(Hex(ls_buf(xy + 2)))) & vbCrLf
'                tOutStr = tOutStr + "Max Receive PDU size is " & HexToTwo(CStr(Hex(ls_buf(xy + 3)))) & HexToTwo(CStr(Hex(ls_buf(xy + 4)))) & vbCrLf
'                tOutStr = tOutStr + "Max Send PDU size is" & HexToTwo(CStr(Hex(ls_buf(xy + 5)))) & HexToTwo(CStr(Hex(ls_buf(xy + 6)))) & vbCrLf
'                tOutStr = tOutStr + "DLMS version number is" & HexToTwo(CStr(Hex(ls_buf(xy + 7)))) & vbCrLf
'                tOutStr = tOutStr + "Quality of service is " & HexToTwo(CStr(Hex(ls_buf(xy + 8)))) & vbCrLf
'                tOutStr = tOutStr + "Dedicated key is " & HexToTwo(CStr(Hex(ls_buf(xy + 9)))) & vbCrLf
'            ElseIf attr_index1 = 6 Then lblAttrOut.Caption = "Authentication mechanism name"
'            ElseIf attr_index1 = 7 Then lblAttrOut.Caption = "LLS secret"
'            ElseIf attr_index1 = 8 Then lblAttrOut.Caption = "Association Status"
'            End If
            xy = 0
            lblObjOut.Caption = "Association LN"
            If attr_index1 = 2 Then
                lblAttrOut.Caption = "Object List"
                If ls_buf(xy) = DT_ARRAY Then
                    xy = xy + 1
                    Arr_Count = (ls_buf(xy) * &H100) Or (ls_buf(xy + 1) And &HFF&) 'packing 2 bytes to an integer
                    xy = xy + 2
                    tOutStr = "Number of Objects = " & Arr_Count
                    For ij = 0 To Arr_Count - 1
                        tOutStr = tOutStr + vbCrLf + vbCrLf + "Object No: " + CStr(ij + 1) + vbCrLf + "--------------------"
                        If ls_buf(xy) = DT_STRUCTURE And ls_buf(xy + 1) = 4 Then
                            xy = xy + 2
                            If ls_buf(xy) <> DT_LONG_UNSIGNED Then Exit For
                            xy = xy + 1
                            tOutStr = tOutStr + vbCrLf + "IC: " + HexToTwo(CStr(Hex(ls_buf(xy)))) + HexToTwo(CStr(Hex(ls_buf(xy + 1))))
                            xy = xy + 2
                            If ls_buf(xy) <> DT_UNSIGNED Then Exit For
                            xy = xy + 1
                            tOutStr = tOutStr + "   Version: " + HexToTwo(CStr(Hex(ls_buf(xy))))
                            xy = xy + 1
                            If ls_buf(xy) <> DT_OCTET_STRING Then Exit For
                            xy = xy + 1
                            data_length1 = ls_buf(xy)
                            xy = xy + 1
                            tOutStr = tOutStr + "   Obis: "
                            For abc = 0 To data_length1 - 1
                                tOutStr = tOutStr + HexToTwo(CStr(Hex(ls_buf(xy)))) + " "
                                xy = xy + 1
                            Next abc
                            If ls_buf(xy) <> DT_STRUCTURE Then Exit For
                            xy = xy + 1
                            If ls_buf(xy) <> 2 Then Exit For
                            xy = xy + 1
                            If ls_buf(xy) <> DT_ARRAY Then Exit For
                            xy = xy + 1
                            num_elements_array = ls_buf(xy)
                            xy = xy + 1
                            For abc = 0 To num_elements_array - 1
                                If ls_buf(xy) = DT_STRUCTURE Then
                                    xy = xy + 1
                                    If ls_buf(xy) = 3 Then
                                        xy = xy + 1
                                        If ls_buf(xy) = DT_INTEGER Then
                                            xy = xy + 1
                                            tOutStr = tOutStr + vbCrLf + "Attrib ID: " + HexToTwo(CStr(Hex(ls_buf(xy))))
                                            xy = xy + 1
                                        End If
                                        If ls_buf(xy) = DT_ENUM Then
                                            xy = xy + 1
                                            tOutStr = tOutStr + "   Access Mode: "
                                            Select Case ls_buf(xy)
                                                Case NO_ASSOC
                                                    tOutStr = tOutStr + "No Access"
                                                Case READ_ONLY
                                                    tOutStr = tOutStr + "Read Only"
                                                Case WRITE_ONLY
                                                    tOutStr = tOutStr + "Write Only"
                                                Case READ_WRITE
                                                    tOutStr = tOutStr + "Read & Write"
                                            End Select
                                            xy = xy + 1
                                            If ls_buf(xy) = DT_NULL_DATA Then tOutStr = tOutStr + "   Access Selector: Null Data"
                                            xy = xy + 1
                                        End If
                                    End If
                                End If
                            Next abc
                            If ls_buf(xy) <> DT_ARRAY Then Exit For
                            xy = xy + 1
                            num_elements_array = ls_buf(xy)
                            xy = xy + 1
                            For abc = 0 To num_elements_array - 1
                                If ls_buf(xy) = DT_STRUCTURE Then
                                    xy = xy + 1
                                    If ls_buf(xy) = 2 Then
                                        xy = xy + 1
                                        If ls_buf(xy) = DT_INTEGER Then
                                            xy = xy + 1
                                            tOutStr = tOutStr + vbCrLf + "Method ID: " + HexToTwo(CStr(Hex(ls_buf(xy))))
                                            xy = xy + 1
                                        End If
                                        If ls_buf(xy) = DT_BOOLEAN Then
                                            xy = xy + 1
                                            tOutStr = tOutStr + "   Access Mode: " + HexToTwo(CStr(Hex(ls_buf(xy))))
                                            xy = xy + 1
                                        End If
                                    End If
                                End If
                            Next abc
                        End If
                    Next ij
                    DATA_DISPLAY = 1
                End If
            ElseIf attr_index1 = 3 Then
                tOutStr = tOutStr + vbCrLf + "Client SAP is " + HexToTwo(CStr(Hex(ls_buf(xy))))
                xy = xy + 1
                tOutStr = tOutStr + vbCrLf + "Server SAP is " + HexToTwo(CStr(Hex(ls_buf(xy)))) + " " + HexToTwo(CStr(Hex(ls_buf(xy + 1))))
                DATA_DISPLAY = 1
            ElseIf attr_index1 = 4 Then
                lblAttrOut.Caption = "Application Context Name"
            ElseIf attr_index1 = 5 Then
                tOutStr = tOutStr + vbCrLf + "Conformance Block is " + HexToTwo(CStr(Hex(ls_buf(xy)))) + " " + HexToTwo(CStr(Hex(ls_buf(xy + 1)))) + " " + HexToTwo(CStr(Hex(ls_buf(xy + 2))))
                tOutStr = tOutStr + vbCrLf + "Max Receive PDU size is " + HexToTwo(CStr(Hex(ls_buf(xy + 3)))) + " " + HexToTwo(CStr(Hex(ls_buf(xy + 4))))
                tOutStr = tOutStr + vbCrLf + "Max Send PDU size is " + HexToTwo(CStr(Hex(ls_buf(xy + 5)))) + " " + HexToTwo(CStr(Hex(ls_buf(xy + 6))))
                tOutStr = tOutStr + vbCrLf + "DLMS version number is  " + HexToTwo(CStr(Hex(ls_buf(xy + 7))))
                tOutStr = tOutStr + vbCrLf + "Quality of service is " + HexToTwo(CStr(Hex(ls_buf(xy + 8))))
                tOutStr = tOutStr + vbCrLf + "Dedicated key is " + HexToTwo(CStr(Hex(ls_buf(xy + 9))))
                DATA_DISPLAY = 1
            ElseIf attr_index1 = 6 Then lblAttrOut.Caption = "Authentication Mechanism Name"
            ElseIf attr_index1 = 7 Then lblAttrOut.Caption = "LLS Secret"
            ElseIf attr_index1 = 8 Then lblAttrOut.Caption = "Association Status"
            End If
        Case 19
            lblObjOut.Caption = "IEC Local Port Setup"
            If attr_index1 = 2 Then lblAttrOut.Caption = "Default Mode"
            If attr_index1 = 3 Then lblAttrOut.Caption = "Default Baud"
            If attr_index1 = 4 Then lblAttrOut.Caption = "Propagation baud"
            If attr_index1 = 5 Then lblAttrOut.Caption = "response time"
            If attr_index1 = 6 Then lblAttrOut.Caption = "Device address"
            If attr_index1 = 7 Then lblAttrOut.Caption = "Password 1"
            If attr_index1 = 8 Then lblAttrOut.Caption = "Password 2"
            If attr_index1 = 9 Then lblAttrOut.Caption = "Password 3"
        Case 20
            lblObjOut.Caption = "ACTIVITY CALENDAR"
            Select Case (attr_index1)
                Case 2
                    xy = 0
                    lblAttrOut.Caption = "Calendar Name Active"
                    For ij = 0 To 7
                        tOutStr = tOutStr + HexToTwo(CStr(Hex(ls_buf(xy)))) + " "
                        xy = xy + 1
                    Next ij
                    DATA_DISPLAY = 1
                Case 3
                    lblAttrOut.Caption = "Season Profile Active"
                    tOutStr = "Array of " & HexToTwo(CStr(Hex(Read_Data_Len / 14))) & " Seasons"
                    For ij = 0 To (Read_Data_Len / 14) - 1
                        tOutStr = tOutStr + vbCrLf + vbCrLf + "-----Season " & HexToTwo(CStr(Hex(ij))) & vbCrLf
                        tOutStr = tOutStr + "Season Name is " & HexToTwo(CStr(Hex(ls_buf(xy))))
                        xy = xy + 1
                        tOutStr = tOutStr + vbCrLf + "Season start Date Time "
                        For abc = 0 To 11
                            tOutStr = tOutStr + HexToTwo(CStr(Hex(ls_buf(xy)))) + " "
                            xy = xy + 1
                        Next abc
                        tOutStr = tOutStr + vbCrLf + "Week Name is " & HexToTwo(CStr(Hex(ls_buf(xy))))
                        xy = xy + 1
                    Next ij
                    DATA_DISPLAY = 1

                Case 4
                        lblAttrOut.Caption = "Week Profile Table"
                        tOutStr = "Array of " & HexToTwo(CStr(Hex(Read_Data_Len / 8))) & " Weeks"
                        For ij = 0 To (Read_Data_Len / 8) - 1
                            tOutStr = tOutStr + vbCrLf + "-----Week " & HexToTwo(CStr(Hex(ij))) + "-----"
                            tOutStr = tOutStr + vbCrLf + "Week Name is " & HexToTwo(CStr(Hex(ls_buf(xy))))
                            xy = xy + 1
                            tOutStr = tOutStr + vbCrLf + "Day ID Monday is " & HexToTwo(CStr(Hex(ls_buf(xy))))
                            xy = xy + 1
                            tOutStr = tOutStr + vbCrLf + "Day ID Tuesday is " & HexToTwo(CStr(Hex(ls_buf(xy))))
                            xy = xy + 1
                            tOutStr = tOutStr + vbCrLf + "Day ID Wednesday is " & HexToTwo(CStr(Hex(ls_buf(xy))))
                            xy = xy + 1
                            tOutStr = tOutStr + vbCrLf + "Day ID Thursday is " & HexToTwo(CStr(Hex(ls_buf(xy))))
                            xy = xy + 1
                            tOutStr = tOutStr + vbCrLf + "Day ID Friday is " & HexToTwo(CStr(Hex(ls_buf(xy))))
                            xy = xy + 1
                            tOutStr = tOutStr + vbCrLf + "Day ID Saturday is " & HexToTwo(CStr(Hex(ls_buf(xy))))
                            xy = xy + 1
                            tOutStr = tOutStr + vbCrLf + "Day ID Sunday is " & HexToTwo(CStr(Hex(ls_buf(xy))))
                            xy = xy + 1
                        Next ij
                        DATA_DISPLAY = 1
                Case 5
                        xy = 0
                        lblAttrOut.Caption = "Day Profile Table"
                        tOutStr = "Array of " & HexToTwo(CStr(Hex(Read_Data_Len / 121))) & "  Day Profiles"
                        For ij = 0 To (Read_Data_Len / 121) - 1
                            xy = (ij * 121)
                            tOutStr = tOutStr + vbCrLf + vbCrLf + "-----Day Profile " & HexToTwo(CStr(Hex(ij))) & " -----"
                            tOutStr = tOutStr + vbCrLf + "Day ID is " & HexToTwo(CStr(Hex(ls_buf(xy))))
                            xy = xy + 1
                            For def = 0 To 9
                                tOutStr = tOutStr + vbCrLf + "---Day Profile action " & HexToTwo(CStr(Hex(def))) & "---"
                                tOutStr = tOutStr + vbCrLf + "Start Time is " & HexToTwo(CStr(Hex(ls_buf(xy)))) & " " & HexToTwo(CStr(Hex(ls_buf(xy + 1)))) & " " & HexToTwo(CStr(Hex(ls_buf(xy + 2)))) & " " & HexToTwo(CStr(Hex(ls_buf(xy + 3))))
                                tOutStr = tOutStr + vbCrLf + "Logical Name is " & HexToTwo(CStr(Hex(ls_buf(xy + 4)))) & " " & HexToTwo(CStr(Hex(ls_buf(xy + 5)))) & " " & HexToTwo(CStr(Hex(ls_buf(xy + 6)))) & " " & HexToTwo(CStr(Hex(ls_buf(xy + 7)))) & " " & HexToTwo(CStr(Hex(ls_buf(xy + 8)))) & " " & HexToTwo(CStr(Hex(ls_buf(xy + 9))))
                                tOutStr = tOutStr + vbCrLf + "Script selector is " & HexToTwo(CStr(Hex(ls_buf(xy + 10)))) & " " & HexToTwo(CStr(Hex(ls_buf(xy + 11)))) + vbCrLf
                                xy = xy + 12
                            Next def
                        Next ij
                        DATA_DISPLAY = 1

                Case 6
                    xy = 0
                    lblAttrOut.Caption = "Calendar Name Passive"
                    For ij = 0 To 7
                        tOutStr = tOutStr + HexToTwo(CStr(Hex(ls_buf(xy)))) + " "
                        xy = xy + 1
                    Next ij
                    DATA_DISPLAY = 1
                Case 7
                        lblAttrOut.Caption = "Season Profile Passive"
                        tOutStr = vbCrLf + "Array of " & HexToTwo(CStr(Hex(Read_Data_Len / 14))) & " Seasons"
                        For ij = 0 To (Read_Data_Len / 14) - 1
                            tOutStr = tOutStr + vbCrLf + "-----Season" & HexToTwo(CStr(Hex(ij))) & " - ----"
                            tOutStr = tOutStr + vbCrLf + "Season Name is " & HexToTwo(CStr(Hex(ls_buf(xy))))
                            xy = xy + 1
                            tOutStr = tOutStr + vbCrLf + "Season start Date Time is "
                            For abc = 0 To 11
                                tOutStr = tOutStr + HexToTwo(CStr(Hex(ls_buf(xy)))) + " "
                                xy = xy + 1
                            Next abc
                            tOutStr = tOutStr + vbCrLf + "Week Name is " & HexToTwo(CStr(Hex(ls_buf(xy))))
                            xy = xy + 1
                        Next ij
                        DATA_DISPLAY = 1
                Case 8
                        lblAttrOut.Caption = "Week Profile Table Passive"
                        tOutStr = "Array of " & HexToTwo(CStr(Hex(Read_Data_Len / 8))) & " Weeks"
                        For ij = 0 To (Read_Data_Len / 8) - 1
                            tOutStr = tOutStr + vbCrLf + "-----Week " & HexToTwo(CStr(Hex(ij))) + "-----"
                            tOutStr = tOutStr + vbCrLf + "Week Name is " & HexToTwo(CStr(Hex(ls_buf(xy))))
                            xy = xy + 1
                            tOutStr = tOutStr + vbCrLf + "Day ID Monday is " & HexToTwo(CStr(Hex(ls_buf(xy))))
                            xy = xy + 1
                            tOutStr = tOutStr + vbCrLf + "Day ID Tuesday is " & HexToTwo(CStr(Hex(ls_buf(xy))))
                            xy = xy + 1
                            tOutStr = tOutStr + vbCrLf + "Day ID Wednesday is " & HexToTwo(CStr(Hex(ls_buf(xy))))
                            xy = xy + 1
                            tOutStr = tOutStr + vbCrLf + "Day ID Thursday is " & HexToTwo(CStr(Hex(ls_buf(xy))))
                            xy = xy + 1
                            tOutStr = tOutStr + vbCrLf + "Day ID Friday is " & HexToTwo(CStr(Hex(ls_buf(xy))))
                            xy = xy + 1
                            tOutStr = tOutStr + vbCrLf + "Day ID Saturday is " & HexToTwo(CStr(Hex(ls_buf(xy))))
                            xy = xy + 1
                            tOutStr = tOutStr + vbCrLf + "Day ID Sunday is " & HexToTwo(CStr(Hex(ls_buf(xy))))
                            xy = xy + 1
                        Next ij
                        DATA_DISPLAY = 1

                Case 9
                        xy = 0
                        lblAttrOut.Caption = "Day Profile Table Passive"
                        'tOutStr = "Read data length  " + Read_Data_Len + vbCrLf
                        tOutStr = "Array of " & HexToTwo(CStr(Hex(Read_Data_Len / 121))) & "  Day Profiles"
                        For ij = 0 To (Read_Data_Len / 121) - 1
                            xy = (ij * 121)
                            tOutStr = tOutStr + vbCrLf + vbCrLf + "-----Day Profile " & HexToTwo(CStr(Hex(ij))) & " -----"
                            tOutStr = tOutStr + vbCrLf + "Day ID is " & HexToTwo(CStr(Hex(ls_buf(xy))))
                            xy = xy + 1
                            For def = 0 To 9
                                tOutStr = tOutStr + vbCrLf + "---Day Profile action " & HexToTwo(CStr(Hex(def))) & "---"
                                tOutStr = tOutStr + vbCrLf + "Start Time is " & HexToTwo(CStr(Hex(ls_buf(xy)))) & " " & HexToTwo(CStr(Hex(ls_buf(xy + 1)))) & " " & HexToTwo(CStr(Hex(ls_buf(xy + 2)))) & " " & HexToTwo(CStr(Hex(ls_buf(xy + 3))))
                                tOutStr = tOutStr + vbCrLf + "Logical Name is " & HexToTwo(CStr(Hex(ls_buf(xy + 4)))) & " " & HexToTwo(CStr(Hex(ls_buf(xy + 5)))) & " " & HexToTwo(CStr(Hex(ls_buf(xy + 6)))) & " " & HexToTwo(CStr(Hex(ls_buf(xy + 7)))) & " " & HexToTwo(CStr(Hex(ls_buf(xy + 8)))) & " " & HexToTwo(CStr(Hex(ls_buf(xy + 9))))
                                tOutStr = tOutStr + vbCrLf + "Script selector is " & HexToTwo(CStr(Hex(ls_buf(xy + 10)))) & " " & HexToTwo(CStr(Hex(ls_buf(xy + 11)))) + vbCrLf
                                xy = xy + 12
                            Next def
                        Next ij
                        DATA_DISPLAY = 1

                Case 10
                        lblAttrOut.Caption = "Activate Passive Calendar time"

                End Select

        Case 22
            xy = 0
            lblObjOut.Caption = "SINGLE ACTION SHEDULE"
            If attr_index1 = 2 Then
                lblAttrOut.Caption = "Executed script"
                tOutStr = "Script Logical name is "
                For ij = 0 To 5
                    tOutStr = tOutStr + HexToTwo(CStr(Hex(ls_buf(xy)))) + " "
                    xy = xy + 1
                Next ij
                tOutStr = tOutStr + vbCrLf + "Script selector is "
                tOutStr = tOutStr + HexToTwo(CStr(Hex(ls_buf(xy))))
                xy = xy + 1
                tOutStr = tOutStr + " " + HexToTwo(CStr(Hex(ls_buf(xy))))
                xy = xy + 1
                DATA_DISPLAY = 1
            ElseIf attr_index1 = 3 Then
                lblAttrOut.Caption = "Type"
                tOutStr = " " + HexToTwo(CStr(Hex(ls_buf(0))))
                DATA_DISPLAY = 1
            ElseIf attr_index1 = 4 Then
                tOutStr = tOutStr + vbCrLf + "Array of " & HexToTwo(CStr(Hex(Read_Data_Len / 9))) & " Execution time"
                For abc = 0 To (Read_Data_Len / 9) - 1
                    tOutStr = tOutStr + vbCrLf + "Execution Time " & HexToTwo(CStr(Hex(abc))) + vbCrLf + "  "
                    For ij = 0 To 3
                        tOutStr = tOutStr + " " + HexToTwo(CStr(Hex(ls_buf(xy))))
                        xy = xy + 1
                    Next ij
                    tOutStr = tOutStr + vbCrLf + "Execution Date " & HexToTwo(CStr(Hex(abc))) + vbCrLf + "  "
                    For ij = 0 To 4
                        tOutStr = tOutStr + " " + HexToTwo(CStr(Hex(ls_buf(xy))))
                        xy = xy + 1
                    Next ij
                Next abc
                DATA_DISPLAY = 1
            End If

        Case 23
            lblObjOut.Caption = "HDLC Setup"
        Case 27
            lblObjOut.Caption = "PSTN Modem Configuration"
        Case 28
            lblObjOut.Caption = "PSTN Auto answer"
    End Select
    
    If attr_index1 = 1 Then
        lblAttrOut.Caption = "Logical name"
    End If
    If DATA_DISPLAY = 0 Then
        For ij = 0 To Read_Data_Len - 1
            tOutStr = tOutStr + HexToTwo(CStr(Hex(ls_buf(ij)))) + " "
        Next ij
    End If
    
    rtxInfoBox.Text = rtxInfoBox.Text + vbCrLf + tOutStr

DO_FINAL_JOB:
    
    Call Enable_ControlButton(True)
    
    lstOBIS.SetFocus
    
End Sub
'*                                                                                                  *

Private Sub imgPowerOff_Click()
    Unload Me
End Sub

Private Sub lstOBIS_Click()
    
    Dim ii_Index As Integer
    
    ii_Index = lstOBIS.ListIndex
    
    txtOBIS(0).Text = sOBIS_Tbl(ii_Index).OBIS_A
    txtOBIS(1).Text = sOBIS_Tbl(ii_Index).OBIS_B
    txtOBIS(2).Text = sOBIS_Tbl(ii_Index).OBIS_C
    txtOBIS(3).Text = sOBIS_Tbl(ii_Index).OBIS_D
    txtOBIS(4).Text = sOBIS_Tbl(ii_Index).OBIS_E
    If IsNumeric(sOBIS_Tbl(ii_Index).OBIS_F) = True Then
        txtOBIS(5).Text = sOBIS_Tbl(ii_Index).OBIS_F
    Else
        If (sOBIS_Tbl(ii_Index).OBIS_F <> "-") And (sOBIS_Tbl(ii_Index).OBIS_F <> "w") Then
            If Already_Read_VZ_Or_Not = False Then
                MsgBox "You must read VZ value before reading data concerned with VZ !", _
                        vbInformation, "Need to read VZ"
            End If
        End If
        Select Case sOBIS_Tbl(ii_Index).OBIS_F
            Case "-":       txtOBIS(5).Text = "-"
            Case "w":       txtOBIS(5).Text = "255"
            Case "VZ":      txtOBIS(5).Text = (gVZ + 100 - 0) Mod 100
            Case "VZ-1":    txtOBIS(5).Text = (gVZ + 100 - 1) Mod 100
            Case "VZ-2":    txtOBIS(5).Text = (gVZ + 100 - 2) Mod 100
            Case Else:      txtOBIS(5).Text = "255"
        End Select
    End If
    txtClassID.Text = sOBIS_Tbl(ii_Index).ClassID
    txtAttrID.Text = sOBIS_Tbl(ii_Index).AttrID
    
    With sOBIS_Tbl(ii_Index)
        If .SetType = "-" Or .SetType = "1" Or .SetType = "2" Then
            cmdWrite_XXX.Enabled = False
        Else
            cmdWrite_XXX.Enabled = True
        End If
    End With
    
End Sub
'*                                                                                                  *

Private Sub lstOBIS_DblClick()

    If Now_Doing_Comm = True Then Exit Sub
    
    Call cmdRead_XXX_Click
    
End Sub
'*                                                                                                  *

'Private Sub rtxInfoBox_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    'If Button = 2 Then
        'PopupMenu mnuDataType
    'End If
'End Sub
'*                                                                                                  *

Private Sub txtOBIS_GotFocus(Index As Integer)
    Call Select_TextBox_Str(txtOBIS(Index))
End Sub
'*                                                                                                  *

Private Sub txtOBIS_LostFocus(Index As Integer)
    If txtOBIS(Index).Text = "-" Then Exit Sub
    If (IsNumeric(txtOBIS(Index).Text) = False) Then
        txtOBIS(Index).Text = "0"
    ElseIf (Val(txtOBIS(Index).Text) < 0) Or (Val(txtOBIS(Index).Text) > 255) Then
        txtOBIS(Index).Text = "0"
    End If
End Sub
'*                                                                                                  *

Private Sub txtOBIS_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 46 Then
        If Index < 5 Then
            txtOBIS(Index + 1).SetFocus
        Else
            txtClassID.SetFocus
        End If
        KeyAscii = 0
        Exit Sub
    End If
    KeyAscii = Check_IntOnly(KeyAscii)
End Sub
'*                                                                                                  *

Private Sub txtAttrID_GotFocus()
    Call Select_TextBox_Str(txtAttrID)
End Sub
'*                                                                                                  *

Private Sub txtClassID_GotFocus()
    Call Select_TextBox_Str(txtClassID)
End Sub
'*                                                                                                  *

Private Sub txtClassID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        txtAttrID.SetFocus
        KeyAscii = 0
        Exit Sub
    End If
    KeyAscii = Check_IntOnly(KeyAscii)
End Sub
'*                                                                                                  *

Private Sub txtAttrID_KeyPress(KeyAscii As Integer)
    KeyAscii = Check_IntOnly(KeyAscii)
End Sub
'*                                                                                                  *

Private Sub Enable_ControlButton(ByVal Enable_Flag As Boolean)

    cmdRead_XXX.Enabled = Enable_Flag
    With sOBIS_Tbl(lstOBIS.ListIndex)
        If .SetType = "-" Or .SetType = "1" Or .SetType = "2" Then
            cmdWrite_XXX.Enabled = False
        Else
            cmdWrite_XXX.Enabled = Enable_Flag
        End If
    End With
    cmdQuit.Enabled = Enable_Flag
    cmdReadVZ.Enabled = Enable_Flag
    cmdCOMset.Enabled = Enable_Flag
    cmdPassword.Enabled = Enable_Flag
    cmdSelectiveLP.Enabled = Enable_Flag
    
    lstOBIS.Enabled = Enable_Flag
    DoEvents
    
    Now_Doing_Comm = Not Enable_Flag
    
End Sub
'*                                                                                                  *

Private Sub Clear_LabelData()

    rtxInfoBox.Text = ""
    lblAttrOut.Caption = ""
    lblObjOut.Caption = ""

End Sub
'*                                                                                                  *

Private Sub Disp_OBIS_Table()

    Dim ii_Index As Integer
    
    lstOBIS.Clear
    
    If OBIS_Table_Total < 2 Then Exit Sub
    
    For ii_Index = 0 To (OBIS_Table_Total - 1) Step 1
        lstOBIS.AddItem sOBIS_Tbl(ii_Index).Descript
    Next ii_Index
    
    lstOBIS.ListIndex = 1
    
End Sub
'*                                                                                                  *

Private Function Check_Support_Set() As Boolean

    Dim ii_Int As Integer
    Dim tmpOBIS As OBIS_Table_Struct
    Dim Tbl_Set_Or_Not As Boolean
    
    With tmpOBIS
        .OBIS_A = txtOBIS(0).Text:  .OBIS_B = txtOBIS(1).Text:  .OBIS_C = txtOBIS(2).Text
        .OBIS_D = txtOBIS(3).Text:  .OBIS_E = txtOBIS(4).Text
        .OBIS_F = Replace(txtOBIS(5).Text, "255", "w")
        .ClassID = txtClassID.Text: .AttrID = txtAttrID.Text
    End With
    
    For ii_Int = 0 To (OBIS_Table_Total - 1) Step 1
        With sOBIS_Tbl(ii_Int)
            If .ClassID = tmpOBIS.ClassID Then
                If .AttrID = tmpOBIS.AttrID Then
                    If (.OBIS_A = tmpOBIS.OBIS_A) And (.OBIS_B = tmpOBIS.OBIS_B) And _
                        (.OBIS_C = tmpOBIS.OBIS_C) And (.OBIS_D = tmpOBIS.OBIS_D) And _
                        (.OBIS_E = tmpOBIS.OBIS_E) And (.OBIS_F = tmpOBIS.OBIS_F) Then
                        If (.SetType <> "-") And (.SetType <> "1") And (.SetType <> "2") Then
                            Check_Support_Set = True
                            lstOBIS.ListIndex = ii_Int
                            Exit For
                        End If
                    End If
                End If
            End If
        End With
    Next ii_Int
    
End Function
'*                                                                                                  *

Private Sub mnuBITSTR_Click()
    Call Copy_Text_And_Paste
    MsgBox Check_Text_And_Calc(4, 1), vbInformation, "BitString Result"
End Sub
'*                                                                                                  *

Private Sub mnuFloat_Click()
    Call Copy_Text_And_Paste
    MsgBox Check_Text_And_Calc(9, 4), vbInformation, "Octet-Float Result"
End Sub
'*                                                                                                  *

Private Sub mnuOctet12_Click()
    Call Copy_Text_And_Paste
    MsgBox Check_Text_And_Calc(9, 12), vbInformation, "Octet-DateTime Result"
End Sub
'*                                                                                                  *

Private Sub mnuSB08_Click()
    Call Copy_Text_And_Paste
    MsgBox Check_Text_And_Calc(15, 1), vbInformation, "Integer8 Result"
End Sub
'*                                                                                                  *

Private Sub mnuSI16_Click()
    Call Copy_Text_And_Paste
    MsgBox Check_Text_And_Calc(16, 2), vbInformation, "Integer16 Result"
End Sub
'*                                                                                                  *

Private Sub mnuUB08_Click()
    Call Copy_Text_And_Paste
    MsgBox Check_Text_And_Calc(17, 1), vbInformation, "Integer16 Result"
End Sub
'*                                                                                                  *

Private Sub mnuUI16_Click()
    Call Copy_Text_And_Paste
    MsgBox Check_Text_And_Calc(18, 2), vbInformation, "Unsigned16 Result"
End Sub
'*                                                                                                  *

Private Sub mnuUL32_Click()
    Call Copy_Text_And_Paste
    MsgBox Check_Text_And_Calc(6, 4), vbInformation, "Unsigned32 Result"
End Sub
'*                                                                                                  *

Private Sub Copy_Text_And_Paste()
    SendKeys "^C":          DoEvents
    txtDataBuf.SetFocus:    DoEvents
    SendKeys "^V":          DoEvents
End Sub
'*                                                                                                  *

Private Function Check_Text_And_Calc(ByVal DataType As Integer, ByVal DataLen As Integer) As String

On Error GoTo ERROR_FOUND

    Dim tSTR As String, tMSG As String
    Dim tARRAY() As String
    Dim tBYTE(3) As Byte
    Dim tDBL As Double, tSNG As Single
    Dim ii_Size As Integer, ii_Int As Integer
    Dim Good_Flag As Boolean
    
    tSTR = txtDataBuf.Text
    If tSTR = "" Then
        tMSG = "Null data selection !" & vbNewLine & "Select exact data in the textbox."
    Else
        tSTR = Replace(tSTR, vbCrLf, "")
        tARRAY = Split(Trim(tSTR), " ")
        ii_Size = UBound(tARRAY)
        If ii_Size <> (DataLen - 1) Then
            tMSG = "Wrong data selection !" & vbNewLine & _
                    "Select exact length of data in the textbox."
        Else
            Good_Flag = True
            For ii_Int = 0 To ii_Size Step 1
                If IsNumeric("&H" & tARRAY(ii_Int)) = False Then
                    Good_Flag = False
                ElseIf Len(tARRAY(ii_Int)) > 2 Then
                    Good_Flag = False
                End If
            Next ii_Int
            If Good_Flag = False Then
                tMSG = "Wrong Hex selection !" & vbNewLine & _
                        "There was wrong Hex data, not numeric."
            Else
                Select Case DataType
                    Case 4      'BitString
                        tMSG = "1-Byte BitString : " & Trim(tSTR) & vbTab & vbNewLine & vbNewLine
                        For ii_Int = 7 To 0 Step -1
                            tMSG = tMSG & "Bit" & CStr(ii_Int) & " - " & _
                                    IIf((CByte("&H" & tARRAY(0)) And (2 ^ ii_Int)) = 0, _
                                    "0", "1") & IIf(ii_Int > 0, vbNewLine, "")
                        Next ii_Int
                    Case 6      'Unsigned32(UL32)
                        tMSG = "4-Byte Unsigned32 : " & Trim(tSTR) & vbTab & vbNewLine & vbNewLine
                        tMSG = tMSG & CStr(Val("&H" & tARRAY(0)) * CDbl(256 ^ 3) + _
                                    Val("&H" & tARRAY(1)) * CDbl(256 ^ 2) + _
                                    Val("&H" & tARRAY(2)) * CDbl(256 ^ 1) + _
                                    Val("&H" & tARRAY(3)))
                    Case 15     'Integer8(SB08)
                        tMSG = "1-Byte Integer8 : " & Trim(tSTR) & vbTab & vbNewLine & vbNewLine
                        ii_Int = CByte("&H" & tARRAY(0))
                        If ii_Int > 127 Then tARRAY(0) = "FF" & HexToTwo(tARRAY(0))
                        tMSG = tMSG & CStr(Val("&H" & tARRAY(0)))
                    Case 16     'Integer16(SI16)
                        tMSG = "2-Byte Integer16 : " & Trim(tSTR) & vbTab & vbNewLine & vbNewLine
                        tMSG = tMSG & CStr(CInt("&H" & tARRAY(0) & HexToTwo(tARRAY(1))))
                    Case 17     'Unsigned8(UB08)
                        tMSG = "1-Byte Unsigned8 : " & Trim(tSTR) & vbTab & vbNewLine & vbNewLine
                        tMSG = tMSG & CStr(CByte("&H" & tARRAY(0)))
                    Case 18     'Unsigned16(UI16)
                        tMSG = "2-Byte Unsigned16 : " & Trim(tSTR) & vbTab & vbNewLine & vbNewLine
                        tMSG = tMSG & CStr(Val("&H" & tARRAY(0)) * 256 + Val("&H" & tARRAY(1)))
                    Case 9      'Octet-String
                        Select Case DataLen
                            Case 4      'Float
                                tMSG = "4-Byte Octet-Float(Revs.) : " & Trim(tSTR) & _
                                        vbTab & vbNewLine & vbNewLine
                                For ii_Int = 0 To 3 Step 1
                                    tBYTE(ii_Int) = CByte("&H" & tARRAY(3 - ii_Int))
                                Next ii_Int
                                CopyMemory tSNG, tBYTE(0), 4
                                tMSG = tMSG & CStr(tSNG)
                            Case 12     'Octet-String(DateTime)
                                tMSG = "12-Byte Octet-DateTime : " & vbNewLine & _
                                        Trim(tSTR) & vbTab & vbNewLine
                                tMSG = tMSG & " (yyyy-mm-dd www, hh:nn:ss.ms)" & vbNewLine & vbNewLine
                                ii_Int = CInt("&H" & tARRAY(0) & HexToTwo(tARRAY(1)))
                                If ii_Int > 0 And ii_Int <= 9999 Then
                                    tMSG = tMSG & Format(ii_Int, "0000") & "-"
                                Else
                                    tMSG = tMSG & "xxxx-"
                                End If
                                ii_Int = CInt("&H" & tARRAY(2))
                                If ii_Int > 0 And ii_Int < 13 Then
                                    tMSG = tMSG & Format(ii_Int, "00") & "-"
                                Else
                                    tMSG = tMSG & "xx-"
                                End If
                                ii_Int = CInt("&H" & tARRAY(3))
                                If ii_Int > 0 And ii_Int < 32 Then
                                    tMSG = tMSG & Format(ii_Int, "00") & " "
                                Else
                                    tMSG = tMSG & "xx "
                                End If
                                ii_Int = CInt("&H" & tARRAY(4))
                                If ii_Int > 0 And ii_Int < 8 Then
                                    tMSG = tMSG & Choose(ii_Int, "Monday", "Tuesday", "Wednesday", _
                                                        "Thursday", "Friday", "Saturday", "Sunday") & ", "
                                Else
                                    tMSG = tMSG & "xxx, "
                                End If
                                ii_Int = CInt("&H" & tARRAY(5))
                                If ii_Int > 0 And ii_Int < 24 Then
                                    tMSG = tMSG & Format(ii_Int, "00") & ":"
                                Else
                                    tMSG = tMSG & "xx:"
                                End If
                                ii_Int = CInt("&H" & tARRAY(6))
                                If ii_Int > 0 And ii_Int < 60 Then
                                    tMSG = tMSG & Format(ii_Int, "00") & ":"
                                Else
                                    tMSG = tMSG & "xx:"
                                End If
                                ii_Int = CInt("&H" & tARRAY(7))
                                If ii_Int > 0 And ii_Int < 60 Then
                                    tMSG = tMSG & Format(ii_Int, "00") & "."
                                Else
                                    tMSG = tMSG & "xx."
                                End If
                                ii_Int = CInt("&H" & tARRAY(8))
                                If ii_Int > 0 And ii_Int < 100 Then
                                    tMSG = tMSG & Format(ii_Int, "00") & vbNewLine
                                Else
                                    tMSG = tMSG & "xx" & vbNewLine
                                End If
                                tMSG = tMSG & "Deviation : " & HexToTwo(tARRAY(9)) & " " & _
                                        HexToTwo(tARRAY(10)) & vbNewLine
                                tMSG = tMSG & "CLK Stat. : " & tARRAY(11)
                            Case Else   'Octet-String(ASCII show)
                                DoEvents
                        End Select
                    Case Else
                    
                End Select
            End If
        End If
    End If
    
    txtDataBuf.Text = ""
    Check_Text_And_Calc = tMSG
    
    Exit Function
    
ERROR_FOUND:

    Check_Text_And_Calc = "Wrong data selection !" & vbNewLine & _
                            "Select exact length of data in the textbox."

End Function
'*                                                                                                  *



