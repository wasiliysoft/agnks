Attribute VB_Name = "Main"
Option Explicit


'������� ��������� ��������� ��������
Public Function Danger() As String
    ' TODO ������������� ������� �����
    ' FIXME ����������� �������� ����������� �������!!!
    If gn������(29).Data = 1 Then
        '���� ������� �� ��. � ���. ������ ������ ������ 5 ��
        If Abs(gnDif(6) - gnDif(2)) <= 0.5 Then
            ROff A1, 223    '������� ���4
            If gbFireTech = True Then
                ROn A1, 24    '� ������� ���3 � ���2
                ROn A0, 2    '������� ���7
            Else
                ROn A1, 16    '� ������� ���3
            End If
        End If
    End If
    frmStart.cmdDanger.Visible = True
End Function



'������� ��������� �����
Public Function ������������() As String
    Dim s           As String
    '������� ��� ����
    ROff A1, 1
    '���� ���, ������� ���2, ������� ���4
    '���� �������������� 20 %
    If (gn������(42).Data = 1) Or (gn������(44).Data = 1) Then
        '���� ���, ������� ���2, ������� ���4
        ROn A1, 42
        ROn A0, 2    '������� ��7
    ElseIf gn������(46).Data = 1 Then    '���� ����� � ����������� ������
        ROn A1, 34    '���� ���, ������� ���4
    End If


    giStage2 = 0
    giStage = 3  '������� �� ���� Danger
    giStage1 = 0
    gbAkkum = False
    frmStart.SSCmdStart.Enabled = False
    gbCmdStart = True
    frmStart.SSCmdStart.Caption = "���� �����"
    ������������ = "������� �����"
End Function

'������� ��������� ���
Public Function ����������() As String
    '���� ������ ���5 - �������
    ' FIXME ����������� �������� ����������� �������!!!
    If gn������(30).Data = 1 Then
        ROff A1, 191
    End If
    '������� ���4
    ROn A1, 32

    giStage2 = 0
    giStage = 1  '������� �� ���� �������
    giStage1 = 1
    giMain������ = 0

    gbAkkum = False
    frmStart.SSCmdStart.Enabled = False
    gbCmdStart = False
End Function
'���� ��������� ��������
Public Function ��������()
    Dim dFullCar    As Double    '����� ���������� �������� � ���� ����������
    Dim s, s1       As String
    Dim MaxIR       As Double    '���������� max ������ ��� �������� ���6
    Dim p           As Double

    ' ������� 8  - �������� ������ �� ��
    If giStage2 = 8 Then
        '�������� �����

        If k4_isOpen Then
            ROff A1, 223 '������� ��4
        Else
            ROff A1, 239 '������� ��3
        End If


        ROn A1, 64      '������� ��5
        giStage2 = 9
        ROn A1, 128      '������� ��6
        gd������1 = 0    '�������� ������ �� ���� ������
        ResetExpenseCounter (2)
        StartOutput (2)
        gbDontStat = True    '������ �������� � ������
        Exit Function
    End If

    '������� 9
    If giStage2 = 9 Then
        If (Abs(gnDif(5) - gnDif(4)) > 0.5) Then
            �������� = "���� �������� "
            '������� ������ �� ���� ������ (�� ����������)
            gdTime = GetTimeCounter(2)
            gd������1 = gd��2
            Exit Function
        Else
            '������� ��������
            '������� ��5
            ROff A1, 191
            ROff A1, 127
            StopOutput (2)
            gbDontStat = False    '����� �������� � ������

            gdTime = GetTimeCounter(2)

            '��������� ���������� �� ��������

            '<<<<���������� ������� ������>>>>
            GMC = GMC + MotorCount
            MotorCount = 0
            StatRS.AddNew
            StatRS("DATA") = Now
            StatRS("GAZ_CAR") = gd������1 / gdPlot    '* 1.42
            StatRS("GAZ_IR1") = gd��1
            StatRS("MOTO") = GMC
                        
            If (Day(gDateRec) < Day(Now)) Or (Month(gDateRec) < Month(Now)) Or (Year(gDateRec) < Year(Now)) Then
                Verify
            End If

            StatRS.Update

            s = Format(Now, "hh:mm:ss") + "        " + Format((gd������1 / gdPlot), "###0.00")
            frmStart.lstStat(0).AddItem s

            gDateRec = Now

            gb�������� = False

            '��������� ��������� �������� ���������� �� ����� �������� �������������
            frmStart.SSCmdStart.Enabled = True
            gbAkkum = True
            giStage = 1    '������� �� ���� ���������
            giStage1 = 0
            giStage2 = 0
            Exit Function

        End If
    End If


    ' ������� 1
    If (giStage2 = 0) And (gbFrmShow = False) Then
        gsMsg = "�������� �������� ?"
        frm������.Show 0
        gbFrmShow = True
        �������� = "������� ������ (��������)"
        Exit Function
    End If

    ' ������� 2
    If (giStage2 = 1) And (gbFrmShow = False) Then
        If giTrigger = 0 Then
            giStage2 = 0
            gb�������� = True
            gbAkkum = False
            giStage = 1    '������� �� ���� ��������
            giStage1 = 1    '����� �� �������� ��� � �����������

            �������� = "������� �� ���� ��������"
            Exit Function
        Else
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
            giStage2 = 3
            gd������1 = 0    '�������� ������ �� ���� ������
            ResetExpenseCounter (2)
            StartOutput (2)
            gbDontStat = True    '������ �������� � ������
            giMain������ = 1
        End If
    End If

    ' ������� 4
    If giStage2 = 3 Then
        dFullCar = gnDif(5)    '���������� �������� � ���� ������
        '������� ������ �������� ����������
        gb�������� = True

        If (gnDif(7) - dFullCar) >= 2 Then    '������� �������� � ������������� � ����
            ROn A1, 128 '������� ��6 - �������� � �� �������������
        End If

        If k4_isOpen Then
            ROff A1, 223 '������� ��4 - �������� ����� ��� ����������
        Else
            ROff A1, 239 '������� ��3 - �������� ����� ��� ����������
        End If

        giStage2 = 4    '��������� � �������� �������� �������������
        ���������
        ���������_1
    End If

    ' ������� 5
    If giStage2 = 4 Then
        '<<<<������� ������ �� ��2>>>>


        MaxIR = GetMassExpense(2)
        If (gbAkkum = False) And ((k6_isOpen And (((MaxIR * 3600) <= gdRashAkkEnd) _
                And (MaxIR > 0)) And (GetTimeCounter(2) >= 5)) Or ((gnDif(7) - gnDif(4)) <= 0.5)) Then           
            ROff A1, 127 '������� ��6
            'Exit Function
        End If
        If (gbAkkum = False) And ((Not (gnDif(4) >= gdUpLevel))) Then
            �������� = "���� �������� "
            '������� ������ �� ���� ������ (�� ����������)
            gd������1 = gd��2
            gdTime = GetTimeCounter(2)
            Exit Function
        ElseIf (gbAkkum = False) Then
            ROff A1, 191 '������� ��5 (��������)
            gbDontStat = False    '����� �������� � ������
            StopOutput (2)
            gdTime = GetTimeCounter(2)

            '��������� ���������� �� ��������
            StatRS.AddNew
            StatRS("DATA") = Now
            StatRS("GAZ_CAR") = gd������1 / gdPlot    '* 1.42

            StatRS("GAZ_IR1") = gd��1
            StatRS("MOTO") = GMC + MotorCount
            GMC = GMC + MotorCount
            MotorCount = 0
            If (Day(gDateRec) < Day(Now)) Or (Month(gDateRec) < Month(Now)) Or (Year(gDateRec) < Year(Now)) Then
                Verify
            End If

            StatRS.Update

            s = Format(Now, "hh:mm:ss") + "        " + Format((gd������1 / gdPlot), "###0.00")
            frmStart.lstStat(0).AddItem s


            gDateRec = Now

            '<<<<���������� ������� ������>>>>
            gb�������� = False            
            ROn A1, 128 '������� ��6 ���������� ������������

            '��������� ��������� �������� ���������� �� ����� �������� �������������
            frmStart.SSCmdStart.Enabled = True
            gbAkkum = True
        End If

        '���������� ������������ �� 200 ��
        If (gnDif(7) < gdUpLevel) And (gbAkkum = True) Then
            �������� = "�������� �������������"
            Exit Function
        Else
            '������� ��6
            ROff A1, 127
        End If
        '���� �������� ����� ������� � �������� ������
        If gbFrmShow = True Then
            frm������.Hide
            frmStart.SSCmdStart.Enabled = True
            gbFrmShow = False
        End If

        '��������� ���������
        s = ����������
        '<<<<���������� ������� ������>>>>
        gb�������� = False


    End If


    ' ������� 7  - �� ����� �������� ������������� ������� �� �������� �����
    If giStage2 = 7 Then       
        ROn A1, 64 '������� ��5
        dFullCar = gnDif(5)    '���������� �������� � ���� ������
        s = "��������� �� �������� �����"
        ResetExpenseCounter (2)
        StartOutput (2)
        ���������
        ���������_1

        giStage2 = 4
        '������� ������ �������� ����������
        giMain������ = 1
        gb�������� = True
        gbAkkum = False
        gd������1 = 0    '�������� ������ �� ���� ������

        gdTime = GetTimeCounter(2)
    End If


    �������� = s
End Function


'�������� ����� � �������� ���������
Public Function �������() As String
    Dim s           As String
    Dim norma       As Boolean
    frmStart.SSCommand2(1).Enabled = True
    frmStart.SSCommand2(0).Enabled = True
    gbFireDVS = False
    gbFireTech = False
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

    If gnDif(14) > 100 Then  '���� �������
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


'�������������� ���������� � �����
Public Function ��������() As String
    If giStage1 = 0 Then
        '���� ���� �������� � �������� ������������ , �� ������� ���4
        If (gnDif(6) - gnDif(2)) >= 0.25 Then
            '������� ��4
            ROn A1, 32
        Else
            '������� ��3 - ��� ������� ���
            ROn A1, 16
        End If
        gbAkkum = True
        giStage1 = 1    ' �������� �� ������ �������
    End If

    If giStage1 = 1 Then
        '������ �������
        If gnDif(14) < 100 Then
            '��� �������� ���
            gbAkkum = True
            frmStart.SSCmdStart.Enabled = True
            '���� ��� ��� ������� � ������ , �� ������� �� ���� �������
            If gbRunDVS = True Then
                giStage2 = 0
                giStage = 0    '������� �� ���� �������
                giStage1 = 0
                gbAkkum = False
                gbRunDVS = False
                frmStart.SSCmdStart.Enabled = False
                gbCmdStart = True
                frmStart.SSCmdStart.Caption = "���� �����"
                'frmStart.Timer2.Enabled = False
                '������� ��� ���
                'TODO ��������� ����������� gn������(25)
                If k2_isOpen Or k3_isOpen Or k4_isOpen Or _
                        k5_isOpen Or k6_isOpen Or k1_isOpen Or _
                        (gn������(25).Data = 1) Then
                    ROff A1, 1
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
        ElseIf gnDif(14) <= 1700 Then
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

Public Sub InitAGNKS()   
    frmStart.tmrMotor.Interval = 65535
    frmStart.tmrMotor.Enabled = False
    

    gbCmdStart = True    '������� ���� �����
    giMain������ = 1    '�������� ��������� � ��������� ��1
    
    InitDisk
    ConnectKKM
    Init_Controllers
    ResetExpenseCounter (1)
    ResetExpenseCounter (2)
End Sub

Private Sub Init_Controllers()
    '������������� ����� ACL8113
    Init_ISO813_Driver
    '������������� ����� Pet48DIO
    Init_DIO_Driver
End Sub




'��������� �������� ��������� ���
Public Function Verify()
    Dim i           As Integer
    Dim j           As Integer
    Dim k           As Integer
    Dim d           As Date
    Dim sum1        As Double
    Dim sum2        As Double
    Dim sum3        As Double
    Dim Old         As String
    Dim s           As String
    Dim s1          As String

    '�������� �������� ����
    d = Now
    sum1 = 0
    sum2 = 0
    sum3 = 0

    If gDateRec < d Then

        For i = 0 To frmStart.lstStat(0).ListCount - 1
            s1 = frmStart.lstStat(0).List(i)
            s = Mid(s1, 17, Len(frmStart.lstStat(0).List(i)) - 1)
            sum1 = sum1 + CDbl(s)
        Next i
        '������ �������� ��������
        Old = s1
        frmStart.lblStat(0).Caption = "�� " + Format(d, "dd")
        frmStart.lblStat(1).Caption = "�� " + Format(d, "mmmm")
        frmStart.lblStat(2).Caption = "�� " + Format(d, "yyyy")

        frmStart.lstStat(0).Clear
        If (Month(gDateRec) < Month(d)) Or ((Month(gDateRec) > Month(d)) And (Year(gDateRec) < Year(d))) Then
            For i = 0 To frmStart.lstStat(1).ListCount - 1
                s1 = frmStart.lstStat(1).List(i)
                s = Mid(s1, 11, Len(frmStart.lstStat(1).List(i)) - 1)
                sum2 = sum2 + CDbl(s)
            Next i
            frmStart.lstStat(1).Clear
            If (sum1 + sum2) <> 0 Then
                s = Format(CStr(Month(gDateRec)), "00") + "        " + Format(CStr(sum2 + sum1), "###0.00")
                frmStart.lstStat(2).AddItem (s)
            End If
        ElseIf (Month(gDateRec) = Month(d)) And (sum1 <> 0) Then
            s = Format(CStr(Day(d - 1)), "00") + "       " + Format(CStr(sum1), "###0.00")
            frmStart.lstStat(1).AddItem (s)
        End If

        If Year(gDateRec) < Year(d) Then
            For i = 0 To frmStart.lstStat(2).ListCount - 1
                s1 = frmStart.lstStat(2).List(i)
                s = Mid(s1, 11, Len(frmStart.lstStat(2).List(i)) - 1)
                sum3 = sum3 + CDbl(s)
            Next i
            frmStart.lstStat(2).Clear
            s = Format(CStr(Year(gDateRec)), "00") + "       " + Format(CStr(sum3), "###0.00")
            frmStart.lstStat(3).AddItem (s)
        End If
        gDateRec = Now

    End If
End Function



Public Function Verify_Damage()
    Dim s           As String
    '������� �������� ��������� ��������
    s = ""
    If gn������(45).Data = 1 Then
        s = s & "����� � ������ ��� ! "
        If gbStopAGNKS = False Then

            '������� ��� ����
            ROff A1, 1
            ROff A0, 0
            '���� ���
            ROn A1, 2
            gbFireDVS = True
            giStage2 = 0
            giStage = 3    '������� �� ���� Danger
            giStage1 = 0
            gbAkkum = False
            frmStart.SSCmdStart.Enabled = False
            gbCmdStart = True
            frmStart.SSCmdStart.Caption = "���� �����"
            frmStart.SSCmdStart.Visible = True
            StopOutput (2)
            gbStopAGNKS = True
        End If
    End If

    If gn������(46).Data = 1 Then
        s = s & "����� � ����. ������ ! "

        If gbStopAGNKS = False Then
            gbFireTech = True
            s = ������������
            gbStopAGNKS = True
            StopOutput (2)
        End If

    End If

    If gn������(42).Data = 1 Then
        s = s & "�������������� 20%(����� ���) ! "
        If gbStopAGNKS = False Then
            s = ������������
            gbStopAGNKS = True
            StopOutput (2)
        End If
    End If
    If gn������(44).Data = 1 Then
        s = s & "�������������� 20%(����.�����) ! "
        If gbStopAGNKS = False Then
            s = ������������
            gbStopAGNKS = True
            StopOutput (2)
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




