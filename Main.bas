Attribute VB_Name = "Main"
Option Explicit


Public Sub InitAGNKS()
    gbCmdStart = True    '������� ���� �����
    InitDisk
    ConnectKKM
    Init_Controllers
    ResetExpenseCounter_2
End Sub

Private Sub Init_Controllers()
    '������������� ����� ACL8113
    Init_ISO813_Driver
    '������������� ����� Pet48DIO
    Init_DIO_Driver
End Sub

' ����� � �������� ���������
' ����������� ��� giStage = 0
Public Function �������() As String
    ������� = ""
    Dim s           As String
    Dim norma       As Boolean
    frmStart.cmdSTOP(1).Enabled = True
    frmStart.cmdSTOP(0).Enabled = True

    s = ""
    norma = True
    gbRunDVS = False
    ' TODO ��������� ����������� ����, �������� k7
    '������� ���� �������� (���� A0 � A1) ? - ����������
    If k2_isOpen Or k3_isOpen Or k4_isOpen Or _
            k5_isOpen Or k6_isOpen Or k1_isOpen Then
        s = "���� �������� ���� !!!"
        norma = False
    End If

    ' FIXME ����������� �������� ����������� �������!!!
    ' �� ��� ��� ������� ��� �� ���������
    If (gn������(25).Data = 1) Then
        s = s & "������ ������� ��� !!!"
        norma = False
    End If

    ' If gnDif(1) <= 0.3 Then   '0.3 ������ ��� ������� ���������
    '   s = s & "��� ���� !!!"
    '   norma = False
    ' End If

    If isClutchOn Then
        s = s & "�������� ����� !!!"
        norma = False
    End If

    If getDVS_RPM > 100 Then  '���� �������
        s = s & "���� ������� ��� !!!"
        norma = False
    End If

    If norma Then
        s = "����� � �������� ��������� ."
        frmStart.SSCmdStart.Enabled = True
        gbOnlyAkk = True
    Else
        s = s & "����� �� ������ !!!"
        frmStart.SSCmdStart.Enabled = False
    End If
    ������� = s
End Function


' ���������� ����������� � �����
' ����������� ��� giStage = 1
Public Function ��������() As String
    �������� = ""
    If giStage1 = 0 Then
        '���� ���� �������� � �������� ������������ , �� ������� ���4
        If (gnDif(6) - gnDif(2)) >= 0.25 Then
            ROn A1, 32 '������� �4
        Else
            ROn A1, 16 '������� �3 - ��� ������� ���
        End If
        gbAkkum = True
        giStage1 = 1    ' �������� �� ������ �������
    End If

    If giStage1 = 1 Then
        '������ �������
        If getDVS_RPM < 100 Then
            '��� �������� ���
            gbAkkum = True
            frmStart.SSCmdStart.Enabled = True
            '���� ��� ��� ������� � ������ , �� ������� �� ���� �������
            If gbRunDVS = True Then
                toStage_0
                gbRunDVS = False

                'frmStart.Timer2.Enabled = False
                'TODO ��������� ����������� gn������(25)
                If k2_isOpen Or k3_isOpen Or k4_isOpen Or _
                        k5_isOpen Or k6_isOpen Or k1_isOpen Or _
                        (gn������(25).Data = 1) Then
                    ROff A1, 1 '������� � 1-6
                End If

            End If

            If isClutchOn Then
                �������� = "��������� �� ����� � ������� !!! �������� ����� "
                Exit Function
            Else
                �������� = "��������� ����� � ������� !!!"
                Exit Function
            End If
            '���� ������� ���
        ElseIf Not (isClutchOn) Then
            �������� = "��������� �� �������� ���� !!!"
            frmStart.SSCmdStart.Enabled = False
            gbOnlyAkk = False
            gbAkkum = False
            Exit Function
        ElseIf getDVS_RPM <= 1700 Then
            �������� = "��� �� ����� �� ������� ����� !!!"
            frmStart.SSCmdStart.Enabled = False
            gbRunDVS = True
            gbOnlyAkk = False
            gbAkkum = False
            Exit Function
        Else
            �������� = "���������� � ������, ����� ���������� !!!"
            frmStart.SSCmdStart.Enabled = True
            gbOnlyAkk = False
            gbRunDVS = True
            giStage2 = 0
            gbAkkum = False
        End If
    End If
End Function

' ����������� ��� giStage = 2
Public Function ��������() As String
    �������� = ""
    Dim s, s1       As String
    Dim MaxIR       As Double    '���������� max ������ ��� �������� ���6
    Dim p           As Double

    '���� �������� ������ �� ��� � ������ ���, �� �� �������
    If (getDVS_RPM < 100) And (gbOnlyAkk = False) Then
        giDVS = giDVS + 1
        If (giDVS > 5) Then
            ROff A1, 0    '������� � 1-6, ���� ���� 2
            ROn A1, 6 ' (2 + 4) ���� ���, ������� �1
            toStage_0
            giDVS = 0
            gbRunDVS = False
            Exit Function
        End If
    Else
        giDVS = 0
    End If
    


    ' ������� 8  - �������� ������ �� ��
    If giStage2 = 8 Then
        '�������� �����
        If k4_isOpen Then
            ROff A1, 223 '������� �4
        Else
            ROff A1, 239 '������� �3
        End If

        ROn A1, 64 '������� ��5
        ROn A1, 128 '������� ��6
        ResetExpenseCounter_2
        StartOutput (2)
        gbDontStat = True    '������ �������� � ������
        giStage2 = 9
        Exit Function
    End If

    '������� 9
    If giStage2 = 9 Then
        If (Abs(gnDif(5) - gnDif(4)) > 0.5) Then
            �������� = "���� �������� "
            Exit Function
        Else
            ROff A1, 191 '������� �5 (��������)
            ROff A1, 127 '������� �6 (���)
            StopOutput (2)
            gbDontStat = False    '����� �������� � ������
            '��������� ���������� �� ��������
            StatRS_Insert

            '��������� ��������� �������� ���������� �� ����� �������� �������������
            frmStart.SSCmdStart.Enabled = True
            gbAkkum = True
            giStage = 1  '������� �� ���� ��������()
            giStage1 = 0
            giStage2 = 0
            Exit Function
        End If
    End If


    ' ������� 1
    If (giStage2 = 0) And (gbFrmShow = False) Then
        frm������.Show vbModeless
        gbFrmShow = True
        �������� = "������� ������ (��������)"
        Exit Function
    End If

    ' ������� 2
    If (giStage2 = 1) And (gbFrmShow = False) Then
        If giTrigger = 0 Then ' �������� �� ��������
            giStage2 = 0
            gbAkkum = False
            giStage = 1  '������� �� ���� ��������()
            giStage1 = 1    '����� �� �������� ��� � �����������

            �������� = "������� �� ���� ��������"
            Exit Function
        Else ' �������� ��������
            giStage2 = 2
            Car = 1
            s = "�������� �����"
        End If
    End If

    ' ������� 3
    If giStage2 = 2 Then
        If Car = 1 Then
            '�������� �����
            ROn A1, 64  '������� ��5
            ResetExpenseCounter_2
            StartOutput (2)
            gbDontStat = True    '������ �������� � ������
            giStage2 = 3
        End If
    End If

    ' ������� 4
    If giStage2 = 3 Then
        If (gnDif(7) - gnDif(5)) >= 2 Then  '������� �������� � ��� � ����������
            ROn A1, 128 '������� �6 - �������� � �� �������������
        End If

        If k4_isOpen Then
            ROff A1, 223 '������� �4 - �������� ����� ��� ����������
        Else
            ROff A1, 239 '������� �3 - �������� ����� ��� ����������
        End If

        giStage2 = 4    '��������� � �������� �������� �������������
    End If

    ' ������� 5
    If giStage2 = 4 Then
        MaxIR = GetMassExpense_2
        If (gbAkkum = False) And _ 
            ((k6_isOpen And (((MaxIR * 3600) <= gdRashAkkEnd) And (MaxIR > 0)) And (MaxIR >= 5)) Or _
             ((gnDif(7) - gnDif(4)) <= 0.5)) Then
            '������� �6 ����
            'gbAkkum = False �
            '������� �6 
            '� (������ ����� ��2 * 3600 <= ��������������������� �� ������ 0) � ������ ����� ��2 >= 5
            '��� ��������� � ���� ������ ��� � ��� �� 0.5
            ROff A1, 127 '������� �6 (���)
            'Exit Function
        End If
        
        If (gbAkkum = False) And ((Not (gnDif(4) >= gdUpLevel))) Then
            �������� = "���� �������� "
            Exit Function
        ElseIf (gbAkkum = False) Then
            ROff A1, 191 '������� �5 (��������)
            StopOutput (2)
            ROn A1, 128 '������� ��6 ���������� ������������

            gbDontStat = False    '����� �������� � ������
            StatRS_Insert '��������� ���������� �� ��������
            
            '��������� ��������� �������� ���������� �� ����� �������� �������������
            frmStart.SSCmdStart.Enabled = True
            gbAkkum = True
        End If

        '���������� ������������ �� 200 ��
        If (gnDif(7) < gdUpLevel) And (gbAkkum = True) Then
            �������� = "�������� �������������"
            Exit Function
        Else
            ROff A1, 127 '������� �6 (���)
        End If
        '���� �������� ����� ������� � �������� ������
        If gbFrmShow = True Then
            frm������.Hide
            frmStart.SSCmdStart.Enabled = True
            gbFrmShow = False
        End If

        '���������� ���������� ???
        Debug.Print "��������� �����������"
        ROff A1, 191 '������� �5 (��������) ???
        ROn A1, 32 '������� �4
        giStage2 = 0
        giStage = 1  '������� �� ���� ��������()
        giStage1 = 1
        gbAkkum = False
        frmStart.SSCmdStart.Enabled = False
        gbCmdStart = False
    End If


    ' ������� 7  - �� ����� �������� ������������� ������� �� �������� �����
    If giStage2 = 7 Then
        s = "��������� �� �������� �����"
        ROn A1, 64 '������� ��5
        ResetExpenseCounter_2
        StartOutput (2)
        gbAkkum = False
        giStage2 = 4
    End If
    �������� = s
End Function

' ������� ��������� ��������� ��������
' ����������� ��� giStage = 3
Public Sub Danger()
    ' TODO ������������� ������� �����
    ' ����������� �������� ����������� ������� "������� �4" !!!
    If gn������(29).Data = 1 Then
        '���� ������� �� ��. � ���. ������ ������ ������ 5 ��
        If Abs(gnDif(6) - gnDif(2)) <= 0.5 Then
            ROff A1, 223    '������� �4
            If isFireTech Then
                ROn A1, 24 '(16 + 8) ������� ���3 � ���2
                ROn A0, 2 '������� ���7
            Else
                ROn A1, 16 '������� ���3
            End If
        End If
    End If
    frmStart.cmdDanger.Visible = True
End Sub

'������� ��������� �����
Public Function ������������() As String
    ROff A1, 1 '������� � 1-6
    '���� �������������� 20 %
    If (gn������(42).Data = 1) Or (gn������(44).Data = 1) Then
        ROn A1, 42 '(2+8+32) ���� ���, ������� �2, ������� �4
        ROn A0, 2  '������� �7
    ElseIf isFireTech Then    '���� ����� � ����������� ������
        ROn A1, 34 '(2+32) ���� ���, ������� ���4
    End If

    giStage2 = 0
    giStage = 3  '������� �� ���� Danger
    giStage1 = 0
    gbAkkum = False
    frmStart.SSCmdStart.Enabled = False
    gbCmdStart = True
    StopOutput (2)
    gbStopAGNKS = True
    ������������ = "������� �����"
End Function

Public Sub toStage_0()
    giStage2 = 0
    giStage = 0
    giStage1 = 0
    gbAkkum = False
    gbCmdStart = True
    frmStart.SSCmdStart.Enabled = False
End Sub

Public Function Verify_Damage() As String
    Verify_Damage = ""
    Dim s           As String
    '������� �������� ��������� ��������
    s = ""
    If gn������(45).Data = 1 Then
        s = s & "����� � ������ ��� ! "
        If gbStopAGNKS = False Then
            ROff A1, 1 '������� � 1-6
            ROff A0, 0 '������� �7, ���� ���� 1
            ROn A1, 2  '���� ���
            giStage2 = 0
            giStage = 3    '������� �� ���� Danger
            giStage1 = 0
            gbAkkum = False
            frmStart.SSCmdStart.Enabled = False
            gbCmdStart = True
            StopOutput (2)
            gbStopAGNKS = True
        End If
    End If

    If isFireTech Then
        s = s & "����� � ����. ������ ! "
        If gbStopAGNKS = False Then
            s = ������������
        End If

    End If

    If gn������(42).Data = 1 Then
        s = s & "�������������� 20%(����� ���) ! "
        If gbStopAGNKS = False Then
            s = ������������
        End If
    End If
    If gn������(44).Data = 1 Then
        s = s & "�������������� 20%(����.�����) ! "
        If gbStopAGNKS = False Then
            s = ������������
        End If
    End If
    If gn������(40).Data = 1 Then
        s = s & "������ ���������� 220 � ! "
    End If

    'If gn������(9).Data = 1 Then
    '  s = s & "����� ������� ����� � ���.������ ! "
    'End If
    If gn������(33).Data = 1 Then
        s = s & "������� tC ���.�������� ��� ! "
    End If

    If gn������(35).Data = 1 Then
        s = s & "������� ����. � ������� ������ ��� ! "
    End If

    If gn������(41).Data = 1 Then
        s = s & "�������������� 10%(����� ���) ! "
    End If
    If gn������(43).Data = 1 Then
        s = s & "�������������� 10%(����.�����) ! "
    End If
    If gn������(47).Data = 1 Then
        s = s & "����� ���-10 ! "
    End If

    If gnDif(13) > 60 Then
        s = s & "�������� ����������� �� ������ ����������� ! "
    End If

    Verify_Damage = s
End Function

Public Function isAll_PSecnsor_OK() As Boolean
    Dim i As Integer
    For i = 2 To 7
         If gnDif(i) = -1 Then
            isAll_PSecnsor_OK = False
            Exit Function
         End If
    Next i
    isAll_PSecnsor_OK = True
End Function
