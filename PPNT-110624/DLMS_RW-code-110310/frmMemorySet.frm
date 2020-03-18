VERSION 5.00
Begin VB.Form frmMemorySet 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '���� ����
   Caption         =   "Internal Memory Set"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4530
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   4530
   StartUpPosition =   1  '������ ���
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
      Left            =   210
      TabIndex        =   2
      Top             =   1980
      Width           =   1635
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
      Left            =   2880
      TabIndex        =   1
      Top             =   1980
      Width           =   1455
   End
   Begin VB.TextBox txtData 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1515
      Left            =   210
      MaxLength       =   203
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmMemorySet.frx":0000
      Top             =   240
      Width           =   4095
   End
End
Attribute VB_Name = "frmMemorySet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*                                                                                                  *

Private Sub Form_Load()
    
    Dim ii_Byte As Byte
    Dim tSTR As String
    
    Set_Confirmed_Or_Not = False
    
    tSTR = ""
    For ii_Byte = 0 To 67 Step 1
        tSTR = tSTR & HexToTwo(Hex(0)) & " "
    Next ii_Byte
    tSTR = Trim(tSTR)
    txtData.Text = tSTR
    
End Sub
'*                                                                                                  *

Private Sub cmdCancel_Click()
    Set_Confirmed_Or_Not = False
    Unload Me
End Sub
'*                                                                                                  *

Private Sub cmdSet_Click()

On Error GoTo ERR_FOUND

    Dim tSTR As String
    Dim tARR() As String
    Dim ii_Byte As Byte
    
    txtData.Text = Trim(UCase(txtData.Text))
    DoEvents
    tSTR = txtData.Text
    tARR = Split(tSTR, " ")
    If Len(tSTR) <> 203 Then        '2byte HEX 68 + Space 67�� = 203bytes
        MsgBox "�Է��� �����Ϳ� �̻��� �ֽ��ϴ�." & vbNewLine & _
                "���ڸ� HEX�� 68���� �� ���̿� ��ĭ 1���� �ʿ��մϴ�.", vbExclamation, "�Է� ����"
        Exit Sub
    End If
    If UBound(tARR) <> 67 Then      '68���� �ƴϸ�...
        MsgBox "Space�� ���е� ���ڰ� 68���� �ƴմϴ�!", vbExclamation, "�Է� ����"
        Exit Sub
    End If
    For ii_Byte = 0 To 67 Step 1
        If Len(tARR(ii_Byte)) <> 2 Then
            MsgBox "�����ʹ� ���ڸ� Hex�� �� �ֽʽÿ�.", vbExclamation, "�Է� ����"
            Exit Sub
        End If
    Next ii_Byte
    For ii_Byte = 0 To 67 Step 1
        ls_buf(ii_Byte) = CByte("&H" & tARR(ii_Byte))
    Next ii_Byte
    
    Erase tARR
    
    Set_Confirmed_Or_Not = True
    Unload Me
    
    Exit Sub

ERR_FOUND:

    MsgBox "�Է� �����Ϳ� �̻��� �ֽ��ϴ�!", vbExclamation, "�Է� ����"
    
End Sub
'*                                                                                                  *


