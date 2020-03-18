Attribute VB_Name = "modOBIS_Table"
Option Explicit
'*                                                                                                  *

Public Const OBIS_Table_File_Name As String = "OBIS_Table.dat"
Public Const MAX_OBIS_Table As Integer = 800

Public OBIS_Table_Total As Integer
Public gVZ As Byte

Type OBIS_Table_Struct
    ClassID As String '* 2
    OBIS_A As String '* 3
    OBIS_B As String '* 3
    OBIS_C As String '* 3
    OBIS_D As String '* 3
    OBIS_E As String '* 3
    OBIS_F As String '* 4
    AttrID As String '* 2
    SetType As String
    SetLen As String
    ReadPage As String
    ReadIndex As String
    ReadOpt As String
    Descript As String
End Type

Public sOBIS_Tbl(MAX_OBIS_Table) As OBIS_Table_Struct

'*                                                                                                  *

Public Sub Init_OBIS_Table()

On Error GoTo ERROR_FOUND

    Dim FileNum
    Dim tSTR As String
    Dim ii_Index As Integer
    
    OBIS_Table_Total = 0
    ii_Index = 0
    
    FileNum = FreeFile
    Open App.Path & "\" & OBIS_Table_File_Name For Input As #FileNum
    
        Line Input #FileNum, tSTR       'First 2 lines are not real OBIS table data
        Line Input #FileNum, tSTR
    
        Do While Not (EOF(FileNum))
        
            Line Input #FileNum, tSTR
            If tSTR = "" Then Exit Do
            sOBIS_Tbl(ii_Index).Descript = tSTR
            
            Line Input #FileNum, tSTR
            If Generate_OBIS_Table(tSTR, sOBIS_Tbl(ii_Index)) = False Then GoTo ERROR_FOUND
            
            ii_Index = ii_Index + 1
            
        Loop
    
    Close #FileNum
    
    OBIS_Table_Total = ii_Index
    
    Exit Sub
    
ERROR_FOUND:

    MsgBox "OBIS table structure data file (OBIS_Table.dat) is bad !" & vbNewLine & _
            "Please check the OBIS table structure data file.", vbExclamation, "File Error"
    Reset
    OBIS_Table_Total = 0

End Sub
'*                                                                                                  *

Private Function Generate_OBIS_Table(ByVal tSTR As String, ByRef myOBIS As OBIS_Table_Struct) As Boolean

On Error GoTo ERROR_FOUND

    Dim tPART() As String
    
    tPART = Split(tSTR, vbTab)
    If UBound(tPART) <> 12 Then GoTo ERROR_FOUND
    
    With myOBIS
        .ClassID = tPART(0)
        .AttrID = tPART(1)
        .OBIS_A = tPART(2)
        .OBIS_B = tPART(3)
        .OBIS_C = tPART(4)
        .OBIS_D = tPART(5)
        .OBIS_E = tPART(6)
        .OBIS_F = tPART(7)
        .SetType = tPART(8)
        .SetLen = tPART(9)
        .ReadPage = tPART(10)
        .ReadIndex = tPART(11)
        .ReadOpt = tPART(12)
        If .SetType <> "-" Then .Descript = .Descript & " _w"
    End With
        
    Generate_OBIS_Table = True
    Exit Function

ERROR_FOUND:

    Generate_OBIS_Table = False
    
End Function
'*                                                                                                  *

