Attribute VB_Name = "IO_Analog"
Option Explicit

Private Const gla����� = &H220

'���������� ����������� ��������� � ����������
'��� ������������� ����� �� 10�
Private Const gKv = 1
    
'���������� ����������� ��������� � ��� (�����������)
'��� ������������� �������� 0,448 ���
Private Const gKi = gKv / 0.2
    
'���������� ����������� ��������� ���� � ��������
Private Const gKp = 1.6 / 16
Private Const gKp_1 = 25 / 16

Public ggACL8113(31) As Double   '��������� �������� ����� 8113
Public gnDif(31)    As Double    ' ��� ������������� ��������(� ���� � ���� ������)

' Function of Driver
Private Declare Function ISO813_DriverInit Lib "ISO813.DLL" () As Integer
Declare Sub ISO813_DriverClose Lib "ISO813.DLL" ()

' Function of AD
Private Declare Function ISO813_AD_Float Lib "ISO813.DLL" (ByVal wBase As Integer, ByVal wChannel As Integer, _
        ByVal wGainCode As Integer, ByVal wBipolar As Integer, _
        ByVal wJmp10v As Integer) As Single


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
            'gla����� - I/O port base address
            'i - ����� ������
            '1 - A/D Gain : 0 0~10 V
            '               1 0~5 V
            '0 - Unipolar
            '1 = 10V     0 = 20 V
            f = ISO813_AD_Float(gla�����, i, 1, 0, 1)
        Next j

        ggACL8113(i) = f
    Next i
End Sub

Sub update_gnDif()
    Dim i as Integer
    Dim dTmp As Double

    For i = 2 To 15
        gnDif(i) = ggACL8113(i) * gKi
    Next i
    gnDif(16) = ggACL8113(16) * 11 '((2000 + 200) / 200)

    '�������� ��� ��1.1 � ��1.2
    For i = 2 To 3
        dTmp = (gnDif(i) - 4)
        If dTmp <= 17 And dTmp >= -1 Then
            gnDif(i) = dTmp * gKp
        Else
            gnDif(i) = -1
        End If
    Next i

    For i = 4 To 7
        '�������� (i - 4) , ���� �� ������������� �� -1
        dTmp = (gnDif(i) - 4)
        If dTmp <= 17 And dTmp >= -1 Then
            gnDif(i) = dTmp * gKp_1
            If gnDif(i) < 0 Then
                gnDif(i) = 0
            End If
        Else
            gnDif(i) = -1
        End If
    Next i

    For i = 8 To 15
        '�������� (i - 4) , ���� �� ������������� �� -1
        dTmp = (gnDif(i) - 4)
        If dTmp <= 17 And dTmp >= -1 Then
            Select Case (i)
                Case 8:  gnDif(i) = 200 * ((dTmp + 1) / 18) - 50 ' ��1, ������ �����������
                Case 9:  gnDif(i) = 12.5 * dTmp - 50    ' ��1.1, ������ �����������
                Case 10: gnDif(i) = 6.25 * dTmp - 50    ' ��2, ������ �����������
                Case 11: gnDif(i) = 6.25 * dTmp - 50    ' ��2.1, ������ �����������
                Case 12: gnDif(i) = 6.25 * dTmp - 50    ' ��3, ������ �����������
                Case 13: gnDif(i) = 150 * ((dTmp + 1) / 18) - 50 ' ��4, ������ ����������� �� ������ �����������
                Case 14: gnDif(i) = 200 * dTmp   ' =3200/16 ������ �������� ���
                Case 15: gnDif(i) = 0.1 / 16 * dTmp     ' ��4 ������������
            End Select
        Else
            gnDif(i) = -1
        End If
    Next i
End Sub

Function getDVS_RPM() as Integer
   getDVS_RPM = gnDif(14)
End Function
