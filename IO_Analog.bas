Attribute VB_Name = "IO_Analog"
Option Explicit

Private Const glaАдрес = &H220

' Function of Driver
Private Declare Function ISO813_DriverInit Lib "ISO813.DLL" () As Integer
Declare Sub ISO813_DriverClose Lib "ISO813.DLL" ()

' Function of AD
Private Declare Function ISO813_AD_Float Lib "ISO813.DLL" (ByVal wBase As Integer, ByVal wChannel As Integer, _
        ByVal wGainCode As Integer, ByVal wBipolar As Integer, _
        ByVal wJmp10v As Integer) As Single


Public ggACL8113(31) As Double   'состояние датчиков платы 8113
Public gnDif(31)    As Double    ' Уже пересчитанные значения(с ними и идет работа)


Public Function Init_ISO813_Driver() As String
    Dim i           As Integer
    Dim msg         As String
    i = ISO813_DriverInit()
    Select Case i
        Case 0: msg = "NoError"
        Case 1: msg = "CheckBoardError"
        Case 2: msg = "DriverOpenError"
        Case 3: msg = "DriverNoOpen"
        Case 4: msg = "AdError"
        Case 5: msg = "OtherError"
        Case 6: msg = "GetDriverVersionError"
        Case &HFFFF: msg = "TimeOutError"
    End Select
    If i <> 0 Then
        MsgBox msg, vbExclamation, "Driver ISO813"
    End If
End Function

Sub update_ggACL8113()
    Dim i           As Long
    Dim j           As Integer
    Dim f           As Double

    For i = 0 To 16
        For j = 0 To 3
            'glaАдрес - I/O port base address
            'i - номер канала
            '1 - A/D Gain : 0 0~10 V
            '               1 0~5 V
            '0 - Unipolar
            '1 = 10V     0 = 20 V
            f = ISO813_AD_Float(glaАдрес, i, 1, 0, 1)
        Next j

        ggACL8113(i) = f
    Next i
End Sub
