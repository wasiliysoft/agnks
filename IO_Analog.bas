Attribute VB_Name = "IO_Analog"
Option Explicit

Private Const glaАдрес = &H220

'****** define the error number *******/
Private Const ISO813_NoError = 0
Private Const ISO813_CheckBoardError = 1
Private Const ISO813_DriverOpenError = 2
Private Const ISO813_DriverNoOpen = 3
Private Const ISO813_AdError = 4
Private Const ISO813_OtherError = 5
Private Const ISO813_GetDriverVersionError = 6
Private Const ISO813_TimeOutError = &HFFFF

' Function of Driver
Declare Function ISO813_DriverInit Lib "ISO813.DLL" () As Integer
Declare Sub ISO813_DriverClose Lib "ISO813.DLL" ()

' Function of AD
Declare Function ISO813_AD_Float Lib "ISO813.DLL" (ByVal wBase As Integer, ByVal wChannel As Integer, _
        ByVal wGainCode As Integer, ByVal wBipolar As Integer, _
        ByVal wJmp10v As Integer) As Single


Public ggACL8113(31) As Double   'состояние датчиков платы 8113
Public gnDif(31)    As Double    ' Уже пересчитанные значения(с ними и идет работа)


Public Function Init_ISO813_Driver() As String

    Dim i           As Integer
    glРезультат = ISO813_DriverInit()
    If glРезультат <> ISO813_NoError Then
        i = MsgBox("Can not initial Driver!!!", , "ISO813 Card Error")
    ElseIf glРезультат = 2 Then
        Init_ISO813_Driver = "Driver open error !"
    Else
        Init_ISO813_Driver = "Плата ACL8113 в норме"
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