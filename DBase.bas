Attribute VB_Name = "DBase"
Option Explicit


Public StatDB       As Database
Public SelectRS     As Recordset
Private StatRS       As Recordset
Private StatWS       As Workspace


Sub init_Database()
    Dim sPath As String
    sPath = App.Path + "\BASE.MDB"
    Set StatWS = DBEngine.Workspaces(0)
    On Error GoTo createDB
    Set StatDB = StatWS.OpenDatabase(sPath)
    Set StatRS = StatDB.OpenRecordset("stat", dbOpenTable)
    Exit Sub
createDB:
    Select Case err.Number
    Case 3024
        Debug.Print "��� ���� ������", sPath
        createDB sPath
        Debug.Print "������� ����� ���� ������"
        Resume
    Case Else
        Dim errStr As String
        errStr = err.Description & ". Code = " & err.Number
        Debug.Print errStr
        MsgBox errStr, vbCritical, "init_Database"
    End Select
End Sub

Private Sub createDB(ByVal sPath As String)
    Dim tdfNew As TableDef

    StatWS.CreateDatabase sPath, dbLangCyrillic
    Set StatDB = StatWS.OpenDatabase(sPath)

    Set tdfNew = StatDB.CreateTableDef("stat")
    With tdfNew
        .Fields.Append .CreateField("data", dbDate)
        .Fields.Append .CreateField("gaz_car", dbDouble)
        .Fields.Append .CreateField("gaz_ir1", dbDouble)
        .Fields.Append .CreateField("moto", dbLong)
    End With
    StatDB.TableDefs.Append tdfNew
    Set tdfNew = Nothing
    StatDB.Close
    Set StatDB = Nothing
End Sub

Function getGMC_from_DB() As Long
    Set SelectRS = StatDB.OpenRecordset("SELECT * From stat ORDER BY stat.data DESC")
    If Not SelectRS.EOF Then
        getGMC_from_DB = SelectRS("MOTO")
    Else
        getGMC_from_DB = 0
    End If
    SelectRS.Close
    Set SelectRS = Nothing
End Function

Sub saveGMC_in_DB()
    Set SelectRS = StatDB.OpenRecordset("SELECT * From stat ORDER BY stat.data DESC")
    If Not SelectRS.EOF Then
        SelectRS.Edit
        GMC = GMC + tmrMotorCounter
        tmrMotorCounter = 0
        SelectRS("MOTO") = GMC
        SelectRS.Update
    End If
    SelectRS.Close
    Set SelectRS = Nothing
End Sub


Sub updateGMC()
    If isAuth = False Then Exit Sub

    Dim h As Long: h = 0
    Dim sInput As String
    sInput = InputBox("������� ����� �������� ��������� ��� � �����")
    sInput = Trim(sInput)
    If (Len(sInput) = 0) Then Exit Sub
    On Error Resume Next
        h = CLng(sInput)
        h = h * 60 ' � ���� ��������� �������� � �������
    On Error GoTo 0
    If  h < 0 Then
        MsgBox "�������� ������ ���� ������ ���� ����� 0.", vbExclamation, "������������ ����"
        Exit Sub
    Else
        Set SelectRS = StatDB.OpenRecordset("SELECT * From stat ORDER BY stat.data DESC")
        If Not SelectRS.EOF Then
            SelectRS.Edit
            GMC = h
            tmrMotorCounter = 0
            SelectRS("MOTO") = GMC
            SelectRS.Update
            MsgBox "���������", vbInformation
        Else
            MsgBox "������ ���� ������, ������������� ��������� ��� ����������.", vbInformation
        End If
        SelectRS.Close
        Set SelectRS = Nothing
    End If
End Sub

Sub load_statistic_from_DB()
'
' ��������� ������� "������" ������� �� ����.
'
    Dim i           As Integer

    Dim dateMin     As Date
    Dim dateMax     As Date
    Dim dateNow     As Date: dateNow = Now

    Dim s           As String
    Dim s1          As String

   frmStart.lblStat(1).Caption = "�� " + Format(Now, "mmmm")
   frmStart.lblStat(2).Caption = "�� " + Format(Now, "yyyy") + " ���"

    If StatRS.EOF Then
        '������ ���� ������
        Exit Sub
    End If

    '�������� ����� ������ ����
    Set SelectRS = StatDB.OpenRecordset("select MIN(DATA) from stat ")
    dateMin = SelectRS(0)
    '�������� ����� ������� ����
    Set SelectRS = StatDB.OpenRecordset("select MAX(DATA) from stat ")
    dateMax = SelectRS(0)
    
    '��� ���������� 4 ������� � ����� ��������� ��� ��������� ��������
    '�� ���� ������� � ������ ������� � �� �����������
    frmStart.lstStat(3).Clear
    For i = Year(dateMin) To Year(dateNow) - 1
        s = "#1/1/" + CStr(i) + " 00:00:00#"
        s1 = "#12/31/" + CStr(i) + " 23:59:59#"
        Set SelectRS = StatDB.OpenRecordset("select SUM(GAZ_CAR) from stat where (stat.DATA between " & s & " AND " & s1 & ")")
        If Not IsNull(SelectRS(0)) Then
            s = Format(CStr(i)) + "        " + Format(CStr(SelectRS(0)), "###0.00")
            frmStart.lstStat(3).AddItem (s)
        End If
    Next i

    '������� "�� ���" � ����� ��������� ��� ��������� ��������
    '�� ������ ����� ���� ������� � ������ � �� ����������� ������.
    frmStart.lstStat(2).Clear
    If Year(dateMax) = Year(dateNow) Then
        For i = 1 To Month(dateNow) - 1
            s = "#" & i & "/1/" & Year(dateNow) & " 00:00:00#"
            s1 = "#" & i & "/" & lastDayByMonth(i, Year(dateNow)) & "/" & Year(dateNow) & " 23:59:59#"
            Set SelectRS = StatDB.OpenRecordset("select SUM(GAZ_CAR) from stat where (stat.DATA between " & s & " AND " & s1 & ")")
            If Not IsNull(SelectRS(0)) Then
                s = Format(i, "00") + "        " + Format(SelectRS(0), "###0.00")
                frmStart.lstStat(2).AddItem (s)
            End If
        Next i
    End If

    '������� "�� �����" � ����� ��������� ��� ��������� �������� �� ��� ������� � 1 ����� ������ � �� ����.
    frmStart.lstStat(1).Clear
    If Month(dateMax) = Month(dateNow) Then
        For i = 1 To Day(dateNow) - 1
            s = "#" & Month(dateNow) & "/" & i & "/" & Year(dateNow) & " 00:00:00#"
            s1 = "#" & Month(dateNow) & "/" & (i + 1) & "/" & Year(dateNow) & " 23:59:59#"
            Set SelectRS = StatDB.OpenRecordset("select SUM(GAZ_CAR) from stat where (stat.DATA between " & s & " AND " & s1 & ")")
            If Not IsNull(SelectRS(0)) Then
                s = Format(i, "00") + "        " + Format(SelectRS(0), "###0.00")
                frmStart.lstStat(1).AddItem (s)
            End If
        Next i
    End If

    '������� "�� �������"
    update_Journal_col_1
End Sub

Private Sub update_Journal_col_1()
    Dim s As String, s1 As String
    Dim i As Long
    frmStart.lstStat(0).Clear
    s = Format(Now, "\#mm\/dd\/yyyy 00:00:00\#")
    s1 = Format(Now, "\#mm\/dd\/yyyy 23:59:59\#")
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
End Sub

Sub StatRS_Insert()
    Dim v As Double: v = gd��2 / agnksConfig.plot
    If v > 0.1 Then ' ������ �� ������������ ������� ��������
        GMC = GMC + tmrMotorCounter
        tmrMotorCounter = 0
        StatRS.AddNew
        StatRS("DATA") = Now
        StatRS("GAZ_CAR") = v
        StatRS("GAZ_IR1") = 0
        StatRS("MOTO") = GMC
        StatRS.Update
        update_Journal_col_1
    End If
    ' FIXME ����� ���������� � ���� ����� ��������
    ' ������� "�� �����", "�� ���"
    ' ��� ���� ��������� ����������� ������ 24/7
    ' ���� ���������� ����� ���� ��� ������
    ' �� � ������� ������ ���������
End Sub

Public Function lastDayByMonth(ByVal m, ByVal yyyy) As Integer
    Select Case m
        Case 1, 3, 5, 7, 8, 10, 12
            lastDayByMonth = 31
        Case 2
            Select Case yyyy
                Case 2004, 2008, 2012, 2016, 2020, 2024, 2028, 2032, 2036, 2040
                    lastDayByMonth = 29
                Case Else
                    lastDayByMonth = 28
            End Select
        Case Else
            lastDayByMonth = 30
    End Select
End Function

'Public Type StatRowType
'    dt              As Date
'    IR1             As Double
'    IR2             As Double
'    Motor           As Long
'End Type
