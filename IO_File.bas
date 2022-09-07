Attribute VB_Name = "IO_File"
Option Explicit

Private pAgnks�onfig As Agnks�onfigType

'��������� ��� ����������������� ����� (configFilePath)
Public Type Agnks�onfigType
    PC As Double        ' ����������� �����������
    price As Double     ' ���� ����
    plot As Double      ' ��������� ����
    motoMinute As Long  ' ���������� � �������
    pwd As String * 7   ' ������
End Type

Private Function configFilePath() As String
   configFilePath = App.Path + "\agnks.config"
End Function

Function agnks�onfig() As Agnks�onfigType
   agnks�onfig = pAgnks�onfig
End Function


'������� ������������� ������, ���������� � �����
'TODO ������� �� ��������� ������� ����������
' � ���������
' ����������� �����
' ���� ����
' ��������� ����
' ����� ������ � ��������� ���������� ����� �� ���� ������
Public Function InitDisk() As Integer
    '���������� ��� ������ : 0 -��� ������
   ' On Error Resume Next
   'TODO �������� ����� ������� ������ �������
    frmStart.MousePointer = vbHourglass

    init_Database
    '�������� �������� �� ����
    GMC = getGMC_from_DB

    load_statistic_from_DB 'TODO ������� �� ������� ������������� �����?

    init_agnksConfig
    
    init_SensorDescr_file
    frmStart.MousePointer = vbArrow
    InitDisk = 0
    Exit Function

ErrorHandler:        '���� ���� �����-������ ������ ���������� -1
    InitDisk = -1
    Exit Function
End Function

Private Sub init_agnksConfig()
    Dim fh As Long: fh = FreeFile
    Dim fLen As Long
    On Error Resume Next
        fLen = FileLen(configFilePath)
    On Error GoTo 0
    If fLen = 0 Then
        With pAgnks�onfig
            .motoMinute = -1
            .PC = -1
            .plot = -1
            .price = -1
            .pwd = "LAB"
        End With
        MsgBox "��������� ���� ������������ �����: " & configFilePath, vbExclamation
    Else
        Open configFilePath For Random As fh Len = Len(pAgnks�onfig)
            Get #fh, 1, pAgnks�onfig
        Close #fh
    End If
    frmStart.lblPC.Caption = Format(agnks�onfig.PC, "0.000")
    frmStart.Label_Price.Caption = pAgnks�onfig.price
    frmStart.lbl_gnPlot.Caption = pAgnks�onfig.plot
End Sub

'FIXME ���������� ������ �����
Sub updatePC()
    Dim s As String
    Dim title As String: title = "DANGER - ���������� ������������ ������������"
    s = InputBox("������� ������", title)
    If (s = Trim(pAgnks�onfig.pwd)) Then
        s = InputBox("������� ����������� �����������", title, Format(agnks�onfig.PC, "0.000"))
        If (CDbl(s) > 0) And (CDbl(s) <= 10) Then
            pAgnks�onfig.PC = CDbl(s)
            saveConfig
            MsgBox "����������� ������", vbInformation
        End If
    Else
        MsgBox "������ �� ������", vbCritical
    End If
    init_agnksConfig
End Sub

' TODO implementation
Sub updatePlot()
    Dim d As Double: d = 0
    Dim sInput As String 
    sInput = InputBox("������� ����� �������� ��������� ����", , Format(agnks�onfig.plot, "0.0000"))
    If (Len(sInput) = 0) Then Exit Sub
    On Error Resume Next
        d = CDbl(sInput)
    On Error GoTo 0 
    If  d >= 1 Or d <= 0.5 Then
        Msgbox "������������ ����", vbExclamation
    Else
        pAgnks�onfig.plot = CDbl(d)
        saveConfig
        init_agnksConfig
        MsgBox "���������", vbInformation
    End If
End Sub

' TODO implementation
Sub updatePrice()
    pAgnks�onfig.price = CDbl(11.7)
    saveConfig
    init_agnksConfig
End Sub
' TODO return result
Private Sub saveConfig()
    Dim fh As Long: fh = FreeFile
    Open configFilePath For Random As fh Len = Len(pAgnks�onfig)
        Put #fh, 1, pAgnks�onfig
    Close #fh
End Sub

Sub updatePWD()
    Dim s As String
    Dim s1 As String
    Dim title As String: title = "DANGER - ���������� ������"
    s = InputBox("������� ������� ������", title)
    If (s = Trim(pAgnks�onfig.pwd)) Then
        s = InputBox("������� ����� ������", title)
        If (Len(s) > 0) And (Len(s) <= 7) Then
            s1 = InputBox("��������� ����� ������", title)
            If (s = s1) Then
                pAgnks�onfig.pwd = s1
                saveConfig
                MsgBox "������ ������", vbInformation
            Else
                MsgBox "������ �� ���������", vbCritical
            End If
        End If
    Else
        MsgBox "������ �� ������", vbCritical
    End If
    init_agnksConfig
End Sub

Private Sub init_SensorDescr_file()
    Dim fh As Long: fh = FreeFile
    Dim s As String
    Dim i As Integer
    
    s = App.Path & "\data.txt"
    Dim fLen As Long
    On Error Resume Next
        fLen = FileLen(s)
    On Error GoTo 0
    If fLen = 0 Then
        MsgBox "��������� ���� ������������ �����: " & s, vbExclamation
        Exit Sub
    End If

    Open s For Input Access Read As fh
        Seek #fh, 1
        '���� ��������� � �������� DIO
        For i = 0 To 47
            Line Input #fh, gn������(i).Note
            frmStart.Label2(i).Caption = gn������(i).Note
        Next i
        '���� ��������� � �������� ISO
        For i = 0 To 15
            Line Input #fh, s
            frmStart.Text2(i).Text = s
        Next i
    Close #fh
End Sub
