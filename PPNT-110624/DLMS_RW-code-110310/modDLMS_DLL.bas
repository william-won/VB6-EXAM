Attribute VB_Name = "modDLMS_DLL"
Option Explicit
'*                                                                                                  *

Public ls_buf(1048576) As Byte
Public data_type1 As Byte
Public data_length1 As Byte
Public data_index1 As Byte
Public buffer1() As Byte

Public gClientID As Byte
Public gContext As Byte
Public gConformance As Long
Public gAuthenication_Mech As Byte

Public DT_NULL_DATA As Byte
Public DT_ARRAY  As Byte
Public DT_STRUCTURE As Byte
Public DT_BOOLEAN As Byte
Public DT_BIT_STRING As Byte
Public DT_DOUBLE_LONG As Byte
Public DT_DOUBLE_LONG_UNSIGNED As Byte
Public DT_FLOATING_POINT As Byte
Public DT_OCTET_STRING As Byte
Public DT_VISIBLE_STRING As Byte
Public DT_TIME As Byte
Public DT_BCD As Byte
Public DT_INTEGER As Byte
Public DT_LONG As Byte
Public DT_UNSIGNED As Byte
Public DT_LONG_UNSIGNED As Byte
Public DT_LONG64 As Byte
Public DT_UNSIGNED_LONG64 As Byte
Public DT_ENUM As Byte
Public DT_REAL32 As Byte
Public DT_REAL64 As Byte
Public DT_OBJECT_IDENTIFIER As Byte

Public NO_ASSOC As Byte
Public READ_ONLY As Byte
Public WRITE_ONLY As Byte
Public READ_WRITE As Byte

Type ASSOC_PART
    tclientID As Byte
    tserverID As Long
    device_address As Long
End Type

Type xdlms_context
    Conformance As Long
    max_recv_pdu As Integer
    max_send_pdu As Integer
    dlms_version As Byte
End Type

Type Stat_Assoc
    assoc_partners As ASSOC_PART
    app_context As Byte
    xdlms_context As xdlms_context
    auth_mech  As Byte
    lls_secret As String
End Type

Type OBISCODE
    a As Byte
    b As Byte
    c As Byte
    d As Byte
    e As Byte
    f As Byte
End Type

Public assoc_index As Stat_Assoc

Type SelectiveAccess
    fromEntry As Long
    toEntry As Long
    fromValue As Integer
    toValue As Integer
End Type

Public sAccess As SelectiveAccess
Public Selective_Or_Not As Boolean

Declare Function openPort Lib "DLMSClientDLL1.dll" Alias "_OpenPort@16" (port As Long, baud As Long, parity As Byte, ByVal opticPort1 As Byte) As Long
Declare Function SendSNRM Lib "DLMSClientDLL1.dll" Alias "_SendSNRM@12" (CID As Byte, SID As Long, DevAdd As Long) As Long
Declare Function disconnect Lib "DLMSClientDLL1.dll" Alias "_Disconnect@0" () As Long
Declare Function ClosePort Lib "DLMSClientDLL1.dll" () As Long
Declare Function Associate Lib "DLMSClientDLL1.dll" Alias "_Associate@8" (assoc_index As Stat_Assoc, ls_buf As Byte) As Long
Declare Function read_XXX Lib "DLMSClientDLL1.dll" Alias "_read_XXX@24" (obis As OBISCODE, ByVal attr_index1 As Byte, ByVal class_id1 As Byte, read_data_length As Long, sAccess As SelectiveAccess, ByVal SelAccess As Byte) As Long
Declare Function write_XXX Lib "DLMSClientDLL1.dll" Alias "_write_XXX@28" (obis As OBISCODE, ByVal attr_index1 As Byte, ByVal class_id1 As Byte, buffer1 As Byte, ByVal data_length_set1 As Long, ByVal data_type_set1 As Byte, ByVal data_index_set1 As Long) As Long

'*                                                                                                  *

Public Function PortOpen() As Boolean

    Dim port, baud As Long
    Dim parity As Byte
    
    port = sCommSET.COM_Port
    baud = sCommSET.Baud_Rate
    parity = sCommSET.Parity_Bit
    
    'If openPort(port, baud, parity, 0) = 0 Then
    If openPort(port, baud, parity, sCommSET.Device) = 0 Then
        frmMain.rtxInfoBox.Text = frmMain.rtxInfoBox.Text + "OpenPort Successful..." + vbCrLf
        PortOpen = True
    Else
        frmMain.rtxInfoBox.Text = frmMain.rtxInfoBox.Text + "OpenPort Failed..." + vbCrLf
        PortOpen = False
    End If
    DoEvents
    
End Function
'*                                                                                                  *

Public Function SNRMSend()
    
    Dim clientID As Byte
    Dim ServerID As Long
    Dim Device_Addr As Long
    
    clientID = gClientID
    ServerID = &H1
    Device_Addr = 1
    
    If SendSNRM(clientID, ServerID, Device_Addr) = 0 Then
        frmMain.rtxInfoBox.Text = frmMain.rtxInfoBox.Text + "SNRM Successful..." + vbCrLf
        SNRMSend = 0
    Else
        frmMain.rtxInfoBox.Text = frmMain.rtxInfoBox.Text + "SNRM Failed..." + vbCrLf
        SNRMSend = 1
    End If
    DoEvents
    
End Function
'*                                                                                                  *

Public Function Assoc()
    
    'Dim assoc_index As Stat_Assoc
    
    'Setting Application Layer parameters
    
    assoc_index.assoc_partners.tclientID = gClientID
    assoc_index.assoc_partners.tserverID = &H1
    assoc_index.assoc_partners.device_address = &H1
    assoc_index.app_context = gContext
    assoc_index.xdlms_context.Conformance = gConformance
    assoc_index.xdlms_context.max_recv_pdu = 1024
    assoc_index.xdlms_context.max_send_pdu = 1024
    assoc_index.xdlms_context.dlms_version = 6
    assoc_index.auth_mech = gAuthenication_Mech
    'assoc_index.lls_secret = "1A2B3C4D"
    
    If Associate(assoc_index, ls_buf(0)) = 0 Then
        frmMain.rtxInfoBox.Text = frmMain.rtxInfoBox.Text + "AARQ Successful..." + vbCrLf
        Assoc = 0
    Else
        frmMain.rtxInfoBox.Text = frmMain.rtxInfoBox.Text + "AARQ Failed" + vbCrLf
        Assoc = 1
    End If
    DoEvents
    
End Function
'*                                                                                                  *

Public Sub discon()
    
    If disconnect() = 0 Then
        frmMain.rtxInfoBox.Text = frmMain.rtxInfoBox.Text + "Disconnected" + vbCrLf
    Else
        frmMain.rtxInfoBox.Text = frmMain.rtxInfoBox.Text + "Disconnect Failed" + vbCrLf
    End If
    DoEvents
    
End Sub
'*                                                                                                  *

Public Function find_data_length1() As Byte

    Select Case (data_type1)
        Case 0          'DT_NULL_DATA
            find_data_length1 = 1:      data_length1 = 0
        Case 1          'DT_ARRAY
            find_data_length1 = 0
        Case 2          'DT_STRUCTURE
            find_data_length1 = 0
        Case 3          'DT_BOOLEAN
            find_data_length1 = 1:      data_length1 = 1
        Case 4          'DT_BIT_STRING
            find_data_length1 = 0
        Case 5          'DT_DOUBLE_LONG
            find_data_length1 = 1:      data_length1 = 4
        Case 6          'DT_DOUBLE_LONG_UNSIGNED
            find_data_length1 = 1:      data_length1 = 4
        Case 7          'DT_FLOATING_POINT
            find_data_length1 = 1:      data_length1 = 4
        Case 9          'DT_OCTET_STRING
            find_data_length1 = 0
        Case 10         'DT_VISIBLE_STRING
            find_data_length1 = 0
        Case 11         'DT_TIME
            find_data_length1 = 0
        Case 13         'DT_BCD
            find_data_length1 = 1:      data_length1 = 4
        Case 15         'DT_INTEGER
            find_data_length1 = 1:      data_length1 = 1
        Case 16         'DT_LONG
            find_data_length1 = 1:      data_length1 = 2
        Case 17         'DT_UNSIGNED
            find_data_length1 = 1:      data_length1 = 1
        Case 18         'DT_LONG_UNSIGNED
            find_data_length1 = 2:      data_length1 = 2
        Case 20         'DT_LONG64
            find_data_length1 = 1:      data_length1 = 8
        Case 21         'DT_UNSIGNED_LONG64
            find_data_length1 = 1:      data_length1 = 8
        Case 22         'DT_ENUM
            find_data_length1 = 1:      data_length1 = 1
        Case 23         'DT_REAL32
            find_data_length1 = 1:      data_length1 = 4
        Case 24         'DT_REAL64
            find_data_length1 = 1:      data_length1 = 8
        Case Default
            find_data_length1 = 0
    End Select
    
End Function
'*                                                                                                  *

Public Sub Init_DataType_CONST_Value()

    DT_NULL_DATA = 0
    DT_ARRAY = 1
    DT_STRUCTURE = 2
    DT_BOOLEAN = 3
    DT_BIT_STRING = 4
    DT_DOUBLE_LONG = 5
    DT_DOUBLE_LONG_UNSIGNED = 6
    DT_FLOATING_POINT = 7
    DT_OCTET_STRING = 9
    DT_VISIBLE_STRING = 10
    DT_TIME = 11
    DT_BCD = 13
    DT_INTEGER = 15
    DT_LONG = 16
    DT_UNSIGNED = 17
    DT_LONG_UNSIGNED = 18
    DT_LONG64 = 20
    DT_UNSIGNED_LONG64 = 21
    DT_ENUM = 22
    DT_REAL32 = 23
    DT_REAL64 = 24
    DT_OBJECT_IDENTIFIER = 6

    NO_ASSOC = 0
    READ_ONLY = 1
    WRITE_ONLY = 2
    READ_WRITE = 3

End Sub
'*                                                                                                  *

Public Sub Init_Comm_Setting_Value()
    
    Dim StrDIR As String
    Dim tSTR As String
    Dim FileNum
    
    StrDIR = App.Path & "\CommSet.dat"
    If Dir(StrDIR) = "" Then
        With sCommSET
            .COM_Port = 1
            .Baud_Rate = 9600
            .Parity_Bit = 0
            .Device = 0
            
            FileNum = FreeFile
            Open StrDIR For Output As #FileNum
                Print #FileNum, CStr(.COM_Port)
                Print #FileNum, CStr(.Baud_Rate)
                Print #FileNum, CStr(.Parity_Bit)
                Print #FileNum, CStr(.Device)
            Close #FileNum
        End With
    Else
        With sCommSET
            FileNum = FreeFile
            Open StrDIR For Input As #FileNum
                Line Input #FileNum, tSTR
                .COM_Port = CByte(tSTR)
                Line Input #FileNum, tSTR
                .Baud_Rate = CLng(tSTR)
                Line Input #FileNum, tSTR
                .Parity_Bit = CByte(tSTR)
            On Error Resume Next
                tSTR = "0"
                Line Input #FileNum, tSTR
                .Device = CByte(tSTR)
            Close #FileNum
        End With
    End If
    
'    gClientID = &H10
'    gContext = 1
'    gConformance = &H1819
'    gAuthenication_Mech = 0
    gPASS_Manufacture = "12345678"
    gPASS_Factory = "1A2B3C4D"
    gPASS_Master = "1A2B3C4D"
    
End Sub
'*                                                                                                  *

