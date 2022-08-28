Attribute VB_Name = "IO_Processing"
Option Explicit


Public Sub ���������()
    '��������� ��������� ������ ����� PET-48DIO
    update_gn48DIO

    '����� ������� ����� ACL8113
    update_ggACL8113
End Sub


Public Sub ���������_1()
    Dim p           As Integer
    Dim r           As Integer
    Dim i           As Integer
    Dim j           As Integer
    Dim Temp        As Double
    Dim IR          As Double
    Dim s           As String
    '���������� ��������� ��������, ���������� � ������ Pet48DIO
    s = ""
    For i = 0 To 5
        p = gn48DIO(i)
        s = s & CStr(p) & " "
        For j = 0 To 7
            r = p Mod 2
            If r = 0 Then
                gn������(8 * i + j).Data = 0
            Else
                gn������(8 * i + j).Data = 1
            End If
            p = Int(p / 2)
        Next j
    Next i

    frmStart.lblPC.Caption = s
    '�������� ���������� �������� � ���:

    For i = 2 To 13
        gnDif(i) = gKi * (ggACL8113(i))
    Next i

    gnDif(15) = gKi * (ggACL8113(15))
    gnDif(14) = ggACL8113(14) * gKi_1

    gnDif(giChanel) = ggACL8113(giChanel) * ((2000 + 200) / 200)
    '�������� ���������� �������� � ����������� (������� �������):

    For i = 8 To 13
        '�������� (i - 4) , ���� �� ������������� �� -1
        Temp = (gnDif(i) - 4)
        If (gnDif(i) - 4) <= 17 And (gnDif(i) - 4) >= -1 Then
            Select Case (i)
                Case 8
                    gnDif(i) = 200 * ((Temp + 1) / 18) - 50
                Case 9
                    gnDif(i) = 12.5 * (gnDif(i) - 4) - 50
                Case 12
                    gnDif(i) = 6.25 * (gnDif(i) - 4) - 50
                Case 13
                    gnDif(i) = 150 * ((Temp + 1) / 18) - 50
                Case Else
                    gnDif(i) = 6.25 * (gnDif(i) - 4) - 50
            End Select
        Else
            gnDif(i) = -1
        End If
    Next i


    '�������� ���������� �������� � �������� (� ���):

    For i = 4 To 7
        '�������� (i - 4) , ���� �� ������������� �� -1
        If (gnDif(i) - 4) <= 17 And (gnDif(i) - 4) >= -1 Then
            gnDif(i) = gKp_1 * (gnDif(i) - 4)
            If gnDif(i) < 0 Then
                gnDif(i) = 0
            End If
        Else
            gnDif(i) = -1
        End If
    Next i

    '���������� ��� ������������� ����� �� �������� ?
    If (gnDif(15) - 4) <= 17 And (gnDif(15) - 4) >= -1 Then
        gnDif(15) = 0.1 / 16 * (gnDif(15) - 4)
    Else
        gnDif(15) = -1
    End If


    '�������� ��� ��1.1 � ��1.2
    For i = 2 To 3
        If (gnDif(i) - 4) <= 17 And (gnDif(i) - 4) >= -1 Then
            gnDif(i) = (gnDif(i) - 4) * gKp
        Else
            gnDif(i) = -1
        End If
    Next i

    ' ������ �������� ���
    If (gnDif(14) - 4) <= 17 And (gnDif(14) - 4) >= -1 Then
        gnDif(14) = (gnDif(14) - 4) * 200   ' =3200/16
    Else
        gnDif(14) = -1
    End If
    Call AddSensorsData(2, gnDif(5), gnDif(11), gnDif(4), 1.5, 0.95 * gdK, 0)
    gd��2 = GetMass(2)
    Temp = GetMassExpense(2)
    '������� ������ (�����) �� ��1
    If giMain������ = 1 Then
        Call AddSensorsData(1, gnDif(2), gnDif(9), gnDif(3), 6, 0.95 * gdK, 0)
    Else
        Call AddSensorsData(1, gnDif(2), gnDif(9), gnDif(3), 6, 0.95 * gdK, -(Temp))
    End If
    IR = GetMass(1)
    gd��1 = IR
End Sub



