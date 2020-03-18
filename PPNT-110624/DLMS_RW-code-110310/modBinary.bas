Attribute VB_Name = "modBinary"
Option Explicit

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As _
    Any, source As Any, ByVal bytes As Long)
'*                                                                                                  *

Public Function HexToTwo(ByVal HexStr As String) As String
    If (Len(HexStr) = 1) Then
        HexToTwo = "0" & HexStr
    Else
        HexToTwo = HexStr
    End If
End Function
'*                                                                                                  *

Public Function Find_Last_Day(ByVal iYear As Integer, ByVal iMonth As Integer) As Integer
    Find_Last_Day = Day(DateAdd("d", -1, DateAdd("m", 1, DateSerial(iYear, iMonth, 1))))
End Function
'*                                                                                                  *

Public Sub Select_TextBox_Str(ByRef myTextBox As TextBox)
    myTextBox.SelStart = 0
    myTextBox.SelLength = Len(myTextBox.Text)
End Sub
'*                                                                                                  *

Public Function Check_IntOnly(KeyAscii As Integer) As Integer
    If (KeyAscii > 47) And (KeyAscii < 58) Or (KeyAscii = 8) Or _
            (KeyAscii = 9) Or (KeyAscii = 45) Then      'Numeric, BS, TAB, "-"
        Check_IntOnly = KeyAscii
    Else
        Check_IntOnly = 0
    End If
End Function
'*                                                                                                  *

Public Function Check_UnsignedOnly(KeyAscii As Integer) As Integer
    If (KeyAscii > 47) And (KeyAscii < 58) Or (KeyAscii = 8) Then   'Numeric, BS
        Check_UnsignedOnly = KeyAscii
    Else
        Check_UnsignedOnly = 0
    End If
End Function
'*                                                                                                  *

Public Function Check_FloatOnly(KeyAscii As Integer) As Integer
    If (KeyAscii > 47) And (KeyAscii < 58) Or (KeyAscii = 8) Or _
            (KeyAscii = 45) Or (KeyAscii = 46) Then      'Numeric, BS, "-", "."
        Check_FloatOnly = KeyAscii
    Else
        Check_FloatOnly = 0
    End If
End Function
'*                                                                                                  *

Public Function Check_AsciiOnly(KeyAscii As Integer) As Integer
    If ((KeyAscii >= 32) And (KeyAscii <= 122)) Or (KeyAscii = 8) Or (KeyAscii = 9) Then
        Check_AsciiOnly = KeyAscii
    Else
        Check_AsciiOnly = 0
    End If
End Function
'*                                                                                                  *

