Attribute VB_Name = "modGlobal"
Option Explicit
'*                                                                                                  *

Public Now_Doing_Comm As Boolean
Public Already_Read_VZ_Or_Not As Boolean
Public Max_LP_Index As Long

Type Comm_Setting_TypeDef
    COM_Port As Byte
    Baud_Rate As Long
    Parity_Bit As Byte
    Device As Byte
End Type

Public sCommSET As Comm_Setting_TypeDef

Public gSet_DataType As Byte
Public gSet_DataLen As Byte

Public Set_Confirmed_Or_Not As Boolean     'frmDataSet CommandButton Click Boolean

Public gPASS_Manufacture As String
Public gPASS_Factory As String
Public gPASS_Master As String
'*                                                                                                  *


