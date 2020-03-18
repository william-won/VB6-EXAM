Attribute VB_Name = "Module1"
Option Explicit

Public Const Setting = ",n,8,1"
Public Return_Massage                   '수신된 프로토콜의 종류? ACK? NAK? ERROR?
Public Registor
Public Reg_Card


Public Function BCC_CHECK(command As String) As Integer
    Dim temp As String
    Dim count As Integer
    Dim bytearray(0 To 30) As Integer
    Dim strarray(0 To 30) As String
    Dim inttemp As Integer
    Dim my As Integer
    Dim C As Integer
    temp = command
    
    If (Len(temp) Mod 2 <> 0) Or (Len(temp) = 0) Then
        MsgBox "수신된 데이터에 오류가 있습니다!"
        Exit Function
    End If
    
    For count = 0 To Len(temp) - 1
        strarray(count) = Mid(temp, count + 1, 1)
        my = Asc(strarray(count)) - 55
        If strarray(count) >= "0" And strarray(count) <= "9" Then
            bytearray(count) = Val(strarray(count))
        ElseIf strarray(count) >= "A" And strarray(count) <= "F" Then
            bytearray(count) = my
        Else
            MsgBox "invalid content", vbCritical, "Error"

            Exit Function
        End If
    Next
    
    For count = 0 To Len(temp) - 1 Step 2
        inttemp = bytearray(count) * 16 + bytearray(count + 1)
        C = C Xor inttemp
    Next count
    
    If C = 0 Then      '결과값 리턴!!
        BCC_CHECK = 1  'BCC OK!!
    Else
        BCC_CHECK = 0  'BCC ERROR!!
    End If
End Function


Public Function ConvertTxt2Binary(command As String, f As Form) As Boolean
    Dim temp As String
    Dim count As Integer
    Dim bytenum As Integer
    Dim bytearray(0 To 100) As Integer
    Dim strarray(0 To 100) As String
    Dim inttemp As Integer
    Dim mine() As Byte
    Dim my As Integer
    Dim oddcount As Integer
    
    'temp = UCase(Command)
    temp = command
    
    If (Len(temp) Mod 2 <> 0) Or (Len(temp) = 0) Then
        MsgBox "똑바로 안하냐?"
        ConvertTxt2Binary = False
        Exit Function
    End If
    
    oddcount = 0
    
    For count = 0 To Len(temp) - 1
        strarray(count) = Mid(temp, count + 1, 1)
        my = Asc(strarray(count)) - 55
        If strarray(count) >= "0" And strarray(count) <= "9" Then
            bytearray(count) = Val(strarray(count))
        ElseIf strarray(count) >= "A" And strarray(count) <= "F" Then
            bytearray(count) = my
        Else
            MsgBox "invalid content", vbCritical, "Error"
            ConvertTxt2Binary = False
            Exit Function
        End If
    Next
    
    ReDim mine((count / 2) - 1)
    
    For count = 0 To Len(temp) - 1 Step 2
        inttemp = bytearray(count) * 16 + bytearray(count + 1)
        oddcount = oddcount + 1
        mine(oddcount - 1) = CByte(inttemp)
    Next count
    
    f.MSComm.Output = mine
    
  '  frmMain.Caption = UBound(mine) - LBound(mine) + 1
    ConvertTxt2Binary = True
End Function



