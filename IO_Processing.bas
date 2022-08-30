Attribute VB_Name = "IO_Processing"
Option Explicit

'���������� ����������� ��������� � ����������
'��� ������������� ����� �� 10�
Private Const gKv = 1
    
'���������� ����������� ��������� � ��� (�����������)
'��� ������������� �������� 0,448 ���
Private Const gKi = gKv / 0.2
    
'���������� ����������� ��������� ���� � ��������
Private Const gKp = 1.6 / 16
Private Const gKp_1 = 25 / 16

'����������� ��������� �������� ���������� �� ��������
'Private Const gKn = (2740 + 448) / 448

Public Sub ���������()
    If isDebug Then
        Exit Sub
    End If
    '��������� ��������� ������ ����� PET-48DIO
    update_gn48DIO

    '����� ������� ����� ACL8113
    update_ggACL8113
End Sub


Public Sub ���������_1()
    Dim i           As Integer
    Dim Temp        As Double

    '���������� ��������� ��������, ���������� � ������ Pet48DIO
    update_gn������

    '�������� ���������� �������� � ���:
    For i = 2 To 15
        gnDif(i) = ggACL8113(i) * gKi
    Next i
    gnDif(16) = ggACL8113(16) * 11 '((2000 + 200) / 200)

    '�������� ��� ��1.1 � ��1.2
    For i = 2 To 3
        Temp = (gnDif(i) - 4)
        If Temp <= 17 And Temp >= -1 Then
            gnDif(i) = Temp * gKp
        Else
            gnDif(i) = -1
        End If
    Next i

    For i = 4 To 7
        '�������� (i - 4) , ���� �� ������������� �� -1
        Temp = (gnDif(i) - 4)
        If Temp <= 17 And Temp >= -1 Then
            gnDif(i) = Temp * gKp_1
            If gnDif(i) < 0 Then
                gnDif(i) = 0
            End If
        Else
            gnDif(i) = -1
        End If
    Next i

    For i = 8 To 15
        '�������� (i - 4) , ���� �� ������������� �� -1
        Temp = (gnDif(i) - 4)
        If Temp <= 17 And Temp >= -1 Then
            Select Case (i)
                Case 8:  gnDif(i) = 200 * ((Temp + 1) / 18) - 50 ' ��1, ������ �����������
                Case 9:  gnDif(i) = 12.5 * Temp - 50    ' ��1.1, ������ �����������
                Case 10: gnDif(i) = 6.25 * Temp - 50    ' ��2, ������ �����������
                Case 11: gnDif(i) = 6.25 * Temp - 50    ' ��2.1, ������ �����������
                Case 12: gnDif(i) = 6.25 * Temp - 50    ' ��3, ������ �����������
                Case 13: gnDif(i) = 150 * ((Temp + 1) / 18) - 50 ' ��4, ������ ����������� �� ������ �����������
                Case 14: gnDif(i) = Temp * 200   ' =3200/16 ������ �������� ���
                Case 15: gnDif(i) = 0.1 / 16 * Temp     ' ��4 ������������
            End Select
        Else
            gnDif(i) = -1
        End If
    Next i


    Call AddSensorsData(2, gnDif(5), gnDif(11), gnDif(4), 1.5, 0.95 * gdK, 0)
    if not isDebug Then
        gd��2 = GetMass(2)
    end if
    '������� ������ (�����) �� ��1
    Temp = -(GetMassExpense(2))
    If giMain������ = 1 Then
        Temp = 0
    End If
    Call AddSensorsData(1, gnDif(2), gnDif(9), gnDif(3), 6, 0.95 * gdK, Temp)
    if not isDebug Then
        gd��1 = GetMass(1)
    End If

    If isDebug Then
        ���������_1_debug
    End If
End Sub


Private Sub ���������_1_debug()
   
    If giStage2 = 9 Then
        gd��2 = gd��2 + 0.15
    Else
        gd��2 = 0
    End If
     '0 A0 Output 0-7
'    gn������(0).Data = 0        ' ���������� ���� 1 (��������)
    'gn������(1).Data = 0        ' ���������� �7
     '1 B0 Input 8-15
'    gn������(8).Data = 0        ' ���� 1 (��������)
'    gn������(9).Data = 0        ' ���������������������
'    gn������(10).Data = 0
'    gn������(11).Data = 0
'    gn������(12).Data = 0
'    gn������(13).Data = 0
'    gn������(14).Data = 0
'    gn������(15).Data = 1          ' �������. ���-�
     '2 C0 Input 16-23
'    gn������(16).Data = 0       ' �2 ������   0
'    gn������(17).Data = 0       ' �3 ������   1
'    gn������(18).Data = 0       ' �4 ������   2
'    gn������(19).Data = 0       ' �5 ������   3
'    gn������(20).Data = 0       ' �6 ������   4
'    gn������(21).Data = 0       ' �1 ������   5
'    gn������(22).Data = 0       '             6
'    gn������(23).Data = 0       ' �7 ������   7
    '3 Config address
    '4 A1 Output 24-31
'    gn������(24).Data = 0       ' ���������� ���� 2 (��������)
'    gn������(25).Data = 0       ' ���� ���������
'    gn������(26).Data = 0       ' ������� ��1
'    gn������(27).Data = 0       ' ������� ��2
'    gn������(28).Data = 0       ' ������� ��3
'    gn������(29).Data = 0       ' ������� ��4
'    gn������(30).Data = 0       ' ������� ��5
'    gn������(31).Data = 0       ' ������� ��6
     '5 B1 Input 32-39
'    gn������(32).Data = 0       ' ���� 2(��������)
'    gn������(33).Data = 0       ' ������� tC ����. ���
'    gn������(34).Data = 0       ' ������ ������������
'    gn������(35).Data = 0       ' Pmax ����� ���
'    gn������(36).Data = 0       ' ����� ���������
'    gn������(37).Data = 0       ' ���. ���(����������)
'    gn������(38).Data = 0       ' ����.�����.������.�1
'    gn������(39).Data = 0       ' ����.�����.������.�2
     '6 C1 Input 40-47
'    gn������(40).Data = 0       ' ������ ����������
'    gn������(41).Data = 0       ' ����� 10% (����� ���)
'    gn������(42).Data = 0       ' ����� 20% (����� ���)
'    gn������(43).Data = 0       ' ����� 10% (����.�����)
'    gn������(44).Data = 0       ' ����� 20% ����
'    gn������(45).Data = 0       ' ����� � ������ ���
'    gn������(46).Data = 0       ' ����� � ���.������
'    gn������(47).Data = 0       ' ����� ���-10

'    gnDif(2) = 1000 ' ��1.1
'    gnDif(3) = 1000 ' ��1.2
'    gnDif(4) = 0 ' ��2.1
'    gnDif(5) = 0 ' ��2.2
'    gnDif(6) = 22 ' ��6, ����������
'    gnDif(7) = 20 ' ��8, �����������
'    gnDif(8) = 10 ' ��1, ������ �����������
'    gnDif(9) = 10 ' ��1.1, ������ �����������
'    gnDif(10) = 10 ' ��2, ������ �����������
'    gnDif(11) = 10 ' ��2.1, ������ �����������
'    gnDif(12) = 10 ' ��3, ������ �����������
'    gnDif(13) = 10 ' ��4, ������ ����������� �� ������ �����������
    'gnDif(14) = 0 ' ������� ���
    'gnDif(15) = 230 ' ��4
'    gnDif(16) = 24.4 ' ���������� ���
   ' Debug.Print gnDif(0)
End Sub


