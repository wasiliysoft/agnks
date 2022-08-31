Attribute VB_Name = "Main"
Option Explicit
Public Function Convert_Date(ss As String)
    Dim s2          As String
    Dim i           As Integer
    s2 = "#"
    For i = 0 To Len(ss)
        If Mid(ss, i + 1, 1) = "." Then
            s2 = s2 & "/"
        Else
            s2 = s2 + Mid(ss, i + 1, 1)
        End If
    Next i
    Convert_Date = s2 & "#"
End Function

'������� ��������� ��������� ��������
Public Function Danger() As String
    ' TODO ������������� ������� �����

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

'������� ������������� ������, ���������� � �����
Public Function InitDisk() As Integer
    '���������� ��� ������ : 0 -��� ������
    Dim i           As Integer
    Dim j           As Integer
    Dim j1          As Integer

    Dim k           As Integer
    Dim t           As MyRecType
    Dim temp1       As MyRecType
    Dim temp2       As MyRecType
    Dim d           As Date
    Dim sum         As Double
    Dim idx1, idx2, idx3 As Integer
    '��� ����� ������������ ������������
    Dim descr       As Integer
    Dim sPath       As String
    Dim rec         As pswd
    Dim s           As String
    Dim s1          As String
    Dim v           As Variant
    On Error Resume Next
    '������� ��� �����

    frmStart.MousePointer = vbHourglass
    s = App.Path + "\base.mdb"
    Set StatWS = DBEngine.Workspaces(0)
    Set StatDB = StatWS.OpenDatabase(s)
    Set StatRS = StatDB.OpenRecordset("stat", dbOpenTable)

    If Not StatRS.EOF Then
        '���� �� ������ ���� ������
        ' StatRS.MoveFirst
        '�������� ����� ������ ����
        Set SelectRS = StatDB.OpenRecordset("select MIN(DATA) from stat ")
        temp1.dt = SelectRS(0)
        '�������� ����� ������� ����
        ' StatRS.MoveLast
        Set SelectRS = StatDB.OpenRecordset("select MAX(DATA) from stat ")
        temp2.dt = SelectRS(0)
        s = Convert_Date(Str(Month(temp2.dt)) & "/" & Day(temp2.dt) & "/" & Year(temp2.dt) & " " & Hour(temp2.dt) & ":" & Minute(temp2.dt) & ":" & Second(temp2.dt))

        Set SelectRS = StatDB.OpenRecordset("SELECT * From stat WHERE stat.data=" & s)

        GMC = SelectRS("MOTO")
        gDateRec = Now    ' temp2.dt

        'temp2.Motor = SelectRS("MOTO")
        d = Now

        '��� ���������� 4 ������� � ����� ��������� ��� ��������� �������� �� ���� ������� � ������ ������� � �� �����������
        For i = Year(temp1.dt) To Year(d) - 1
            s = "1/1/" + CStr(i)
            s = Convert_Date(s)
            s1 = "12/31/" + CStr(i)
            s1 = Convert_Date(s1)
            Set SelectRS = StatDB.OpenRecordset("select SUM(GAZ_CAR) from stat where (stat.DATA between " & s & " AND " & s1 & ")")
            If IsNull(SelectRS(0)) Then
            Else
                s = Format(CStr(i)) + "        " + Format(CStr(SelectRS(0)), "###0.00")
                frmStart.lstStat(3).AddItem (s)

            End If
        Next i
        '��� ���������� 3 ������� � ����� ��������� ��� ��������� �������� �� ������ ������� � ������ ����� ���� � �� ����.
        If Year(temp2.dt) = Year(d) Then
            For i = 1 To Month(d) - 1
                s = CStr(i) + "/1/" + CStr(Year(d))
                s = Convert_Date(s)

                Select Case i
                    Case 1, 3, 5, 7, 8, 10, 12
                        s1 = CStr(i) + "/31/" + CStr(Year(d))
                    Case 2
                        Select Case Year(d)
                            Case 2004, 2008, 2012, 2016, 2020, 2024, 2028, 2032, 2036, 2040
                                s1 = CStr(i) + "/29/" + CStr(Year(d))
                            Case Else
                                s1 = CStr(i) + "/28/" + CStr(Year(d))
                        End Select
                    Case Else
                        s1 = CStr(i) + "/30/" + CStr(Year(d))
                End Select

                s1 = Convert_Date(s1)
                Set SelectRS = StatDB.OpenRecordset("select SUM(GAZ_CAR) from stat where (stat.DATA between " & s & " AND " & s1 & ")")
                v = SelectRS(0)
                If IsNull(v) Then
                Else
                    s = Format(CStr(i), "00") + "        " + Format(CStr(SelectRS(0)), "###0.00")
                    frmStart.lstStat(2).AddItem (s)
                End If
            Next i
        End If
        '��� ���������� 2 ������� � ����� ��������� ��� ��������� �������� �� ��� ������� � 1 ����� ������ � �� ����.

        frmStart.lblStat(0).Caption = "�� " + Format(d, "dd")
        frmStart.lblStat(1).Caption = "�� " + Format(d, "mmmm")
        frmStart.lblStat(2).Caption = "�� " + Format(d, "yyyy")

        If Month(temp2.dt) = Month(d) Then
            For i = 1 To Day(d) - 1
                s = CStr(Month(d)) + "/" + CStr(i) + "/" + CStr(Year(d))
                s = Convert_Date(s)
                s1 = CStr(Month(d)) + "/" + CStr(i + 1) + "/" + CStr(Year(d))
                s1 = Convert_Date(s1)

                Set SelectRS = StatDB.OpenRecordset("select SUM(GAZ_CAR) from stat where (stat.DATA between " & s & " AND " & s1 & ")")
                v = SelectRS(0)
                If IsNull(v) Then
                Else
                    s = Format(CStr(i), "00") + "        " + Format(CStr(SelectRS(0)), "###0.00")
                    frmStart.lstStat(1).AddItem (s)
                End If
            Next i
        End If

        '���� ���� ��������� �������� = �������, �� � 1 ������� ������� �������� �� ��� ����.
        s = Format(d, "mm/dd/yyyy")
        s = Convert_Date(s)
        s1 = Format(d + 1, "mm/dd/yyyy")
        s1 = Convert_Date(s1)
        'frmShow.MousePointer = vbHourglass
        Set SelectRS = StatDB.OpenRecordset("select * from stat where DATA between " & s & " AND " & s1)
        If SelectRS.RecordCount >= 1 Then
            SelectRS.MoveLast
            SelectRS.MoveFirst

            For i = 0 To SelectRS.RecordCount - 1
                s = ""
                s = Format(CStr(SelectRS("Data")), "hh:mm:ss") + "        " + Format(CStr(SelectRS("GAZ_CAR")), "###0.00")
                frmStart.lstStat(0).AddItem (s)
                SelectRS.MoveNext
            Next i
        End If

        'GMC ��������� SelectRS("MOTO") ��������� �� ���� ������
        'gDateRec ��������� SelectRS("Data")��������� �� ���� ������
        '
        'd = Now
        'sum = 0
        'k = Int(gdaStat1(0).IR1) '���������� ������� � ������� 1
        'gDateRec = gdaStat1(1).dt
        'GMC = gdaStat1(k).Motor

        'RefreshStat

        'Call showfields
    Else
        GMC = 0
        gDateRec = Now    ' temp2.dt

    End If

    gDateRec = Now
    sPath = "C:\Winnt\dll32.dll"
    descr = FreeFile
    Open sPath For Random As descr Len = Len(rec)
    If FileLen(sPath) = 0 Then
        rec.pwd = "LAB"
        rec.PC = 1
        Put #descr, 1, rec
        Password = "LAB"
        gdK = rec.PC
    Else
        Get #descr, 1, rec
        Password = Trim(rec.pwd)
        gdK = rec.PC
    End If
    Close #descr

    InitDisk = 0


    frmStart.MousePointer = vbArrow
    gdPlot = 0.7
    Dim fh          As Long
    fh = FreeFile
    s = App.Path & "\price.txt"
    Open s For Input Access Read As fh
    Seek #fh, 1
    Line Input #fh, s
    gdPrice = CDbl(s)
    Line Input #fh, s
    gdPlot = CDbl(s)
    Close #fh
    frmStart.Label_Price.Caption = gdPrice
    If gdPlot < 0.5 Then gdPlot = 0.7
    If gdPlot > 1 Then gdPlot = 0.7
    frmStart.Caption = frmStart.Caption & " ��������� ���� = " & CStr(gdPlot) & " ��/�3"
    Exit Function

ErrorHandler:        '���� ���� �����-������ ������ ���������� -1
    InitDisk = -1
    Exit Function

End Function

'��������� ��������� ���������� � ���� ����������
Public Sub RefreshStat()
    Dim i           As Integer
    Dim s           As String
    For i = 0 To 3
        frmStart.lstStat(i).Clear
    Next i

    If gdaStat1(1).dt = 0 Then
        frmStart.lblStat(0).Caption = "�� " & Format(Now, "d")
    Else
        frmStart.lblStat(0).Caption = "�� " & Format(gdaStat1(1).dt, "d")
    End If

    If gdaStat2(1).dt = 0 Then
        frmStart.lblStat(1).Caption = "�� " & Format(Now, "mmmm")
    Else
        frmStart.lblStat(1).Caption = "�� " & Format(gdaStat2(1).dt, "mmmm")
    End If

    If gdaStat3(1).dt = 0 Then
        frmStart.lblStat(2).Caption = "�� " & Format(Now, "yyyy")
    Else
        frmStart.lblStat(2).Caption = "�� " & Format(gdaStat3(1).dt, "yyyy")
    End If


    For i = 1 To gdaStat1(0).IR1
        s = Format(gdaStat1(i).dt, "hh:mm:ss") + "     " + Format(gdaStat1(i).IR2, "###0.00")
        frmStart.lstStat(0).AddItem s
    Next

    For i = 1 To gdaStat2(0).IR1
        s = Format(gdaStat2(i).dt, "mmmm d yyyy") + "     " + Format(gdaStat2(i).IR2, "###0.00")
        frmStart.lstStat(1).AddItem s
    Next

    For i = 1 To gdaStat3(0).IR1
        s = Format(gdaStat3(i).dt, "mmmm ") + "     " + Format(gdaStat3(i).IR2, "###0.00")
        frmStart.lstStat(2).AddItem s
    Next

    For i = 1 To gdaStat4(0).IR1
        s = Format(gdaStat4(i).dt, "yyyy") + "     " + Format(gdaStat4(i).IR2, "###0.00")
        frmStart.lstStat(3).AddItem s
    Next

End Sub


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

        If gn������(18).Data = 1 Then
            '������� ��4
            ROff A1, 223
        Else
            '������� ��3
            ROff A1, 239
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

        If gn������(18).Data = 1 Then
            '������� ��4 - �������� ����� ��� ����������
            ROff A1, 223
        Else
            '������� ��3 - �������� ����� ��� ����������
            ROff A1, 239
        End If

        giStage2 = 4    '��������� � �������� �������� �������������
        ���������
        ���������_1
    End If

    ' ������� 5
    If giStage2 = 4 Then
        '<<<<������� ������ �� ��2>>>>


        MaxIR = GetMassExpense(2)
        If (gbAkkum = False) And (((gn������(20).Data = 1) And (((MaxIR * 3600) <= gdRashAkkEnd) _
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
    '������� ���� �������� (���� A0 � A1) ? - ����������
    If (gn������(16).Data = 1) Or (gn������(17).Data = 1) Or (gn������(18).Data = 1) Or _
            (gn������(19).Data = 1) Or (gn������(20).Data = 1) Or (gn������(21).Data = 1) Then
        s = "���� �������� ���� !!!"
        norma = False
    End If

    If (gn������(25).Data = 1) Then
        s = s & "������ ������� ��� !!!"
        norma = False
    End If

    ' If gnDif(1) <= 0.3 Then   '0.3 ������ ��� ������� ���������
    '   s = s & "��� ���� !!!"
    '   norma = False
    ' End If

    If gn������(36).Data = 1 Then
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
                If (gn������(16).Data = 1) Or (gn������(17).Data = 1) Or (gn������(18).Data = 1) Or _
                        (gn������(19).Data = 1) Or (gn������(20).Data = 1) Or (gn������(21).Data = 1) Or _
                        (gn������(25).Data = 1) Then
                    ROff A1, 1
                End If

            End If

            If gn������(36).Data = 0 Then
                �������� = "��������� ����� � ������� !!!"
                Exit Function
            Else
                �������� = "��������� �� ����� � ������� !!! �������� ����� "
                Exit Function
            End If
            '���� ������� ���
        ElseIf gn������(36).Data = 0 Then
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
    Dim err         As Integer
    
    MotorCount = 0
    frmStart.tmrMotor.Interval = 65535
    frmStart.tmrMotor.Enabled = False

    ' ������������� ����
    giDVS = 0

    giStage = 0  '������������� ��������� �������� (��������� �� �������)
    giStage1 = 0
    giStage2 = 0
    gbFrmShow = False
    frmStart.SSTab1.Tab = 3
    gbCmdStart = True    '������� ���� �����
    gbAkkum = False
    Car = 0
    glAver = 0
    glCounter = 0
    gbRunDVS = False
    gd������1 = 0

    gdRashAkkEnd = 65
    gdK = 1


    gbDontStat = False

    gbStopAGNKS = False
    giMain������ = 1    '�������� ��������� � ��������� ��1
    err = InitDisk()
    If err = -1 Then
        '��������� ������ ������ � ������
        '�������� ��� ������ �� ����
        '���� ���� ������ ��������� � ��������, �� �������� ���
    End If
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

Public Sub ShowPict()
    Dim s           As String
    With frmStart
        '���������� �������
        If gn������(21).Data = 1 Then
            .��1(0).Visible = False
            .��1(1).Visible = True
        Else
            .��1(0).Visible = True
            .��1(1).Visible = False
        End If

        If gn������(16).Data = 1 Then
            .��2(0).Visible = False
            .��2(1).Visible = True
            .�����(0).Visible = True
        Else
            .��2(0).Visible = True
            .��2(1).Visible = False
            .�����(0).Visible = False
        End If

        If gn������(17).Data = 1 Then
            .��3(0).Visible = False
            .��3(1).Visible = True
        Else
            .��3(0).Visible = True
            .��3(1).Visible = False
        End If

        If gn������(18).Data = 1 Then
            .��4(0).Visible = False
            .��4(1).Visible = True
        Else
            .��4(0).Visible = True
            .��4(1).Visible = False
        End If

        If gn������(19).Data = 1 Then
            .��5(0).Visible = False
            .��5(1).Visible = True
        Else
            .��5(0).Visible = True
            .��5(1).Visible = False
        End If

        If gn������(20).Data = 1 Then
            .��6(0).Visible = False
            .��6(1).Visible = True
        Else
            .��6(0).Visible = True
            .��6(1).Visible = False
        End If

        If gn������(23).Data = 1 Then
            .��7(0).Visible = False
            .��7(1).Visible = True
            .�����(1).Visible = True
        Else
            .��7(0).Visible = True
            .��7(1).Visible = False
            .�����(1).Visible = False
        End If



        '���������� ������
        If gn������(36).Data = 1 Then
            .�����.BackColor = &HFF&
        Else
            .�����.BackColor = &HC0C0C0
        End If

        '��������� ����������� ���������� ���
        If gn������(33).Data = 1 Then
        Else
        End If

        '����� � ������ ���
        If gn������(45).Data = 1 Then
        Else
        End If

        '����� � ��������������� ������
        If gn������(46).Data = 1 Then
        Else
        End If

        '��� � ������ ��� 10%
        If (gn������(41).Data = 1) Then
        Else
        End If
        '��� � ������ ��� 20%
        If (gn������(42).Data = 1) Then
        Else
        End If

        '��� � ��������������� ������ 10%
        If (gn������(43).Data = 1) Then
        Else
        End If

        '��� � ��������������� ������ 20%
        If (gn������(44).Data = 1) Then
        Else
        End If
    End With

    ' TODO ������� �� ���� �������
    s = Format(gdK, "0.000")
    s = s + "   - �����������"
    frmStart.lblPC.Caption = s


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
    'If gDateRec < d Then
    '           MsgBox ("Verify" + Str(gDateRec) + Str(d))

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
        ElseIf (Year(gDateRec) = Year(d)) And (sum2 <> 0) Then
            's = Format(CStr(Month(gDateRec)), "00") + "       " + Format(CStr(sum2), "###0.00")

            '    s = Format(CStr(Month(d)), "00") + "       " + Format(CStr(sum2), "###0.00")
            'frmStart.lstStat(2).AddItem (s)
        End If
        gDateRec = Now

    End If
    'End If




    'GMC = temp2.Motor
    ' If Not StatRS.EOF Then
    '    gDateRec = temp2.dt
    ' Else
    '    gDateRec = Now  ' temp2.dt
    ' End If
    ' frmStart.MousePointer = vbArrow
    'd = Now
    'sum = 0
    'k = Int(gdaStat1(0).IR1) '���������� ������� � ������� 1
    'gDateRec = gdaStat1(1).dt
    'GMC = gdaStat1(k).Motor

    'RefreshStat

    'Call showfields
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




