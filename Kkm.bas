Attribute VB_Name = "KKM"
Option Explicit
'ConnectKKM - ����������� ��� �������� �������� �����



Public giErrorKKM   As Integer  '��� ������ ��� ���������� ������� ������ � ��������� ���
Public gsErrorKKM   As String
Public pass         As Long
Public gs��������   As String
Public gi����������� As Integer
Public glpassKKM    As Long     '������ ���������� ��� ��� ������ � ������� ���
Public DrvFR        As Object   '�������� ������� �������� ��

'  Public Drvfr As Object
Public Sub StatusKKM()
    On Error GoTo err
    '������� ������� ������� ��� ���������� � ������ ������ ��������� ����� ��������� ����������
    '� ����� ������ ��� ������� �� ��������� PrintCheckKKM ��� ���������� �������� ������ � �������� ������
    DrvFR.Password = pass     '��������� ������ ��� ����������� ������� ���
    'DrvFR.GetECRStatus '������ ������� ���
    DrvFR.GetShortECRStatus     '������ ������� ���
    giErrorKKM = DrvFR.ResultCode
    gsErrorKKM = DrvFR.ResultCodeDescription
    gs�������� = DrvFR.ECRModeDescription
    gi����������� = DrvFR.ECRAdvancedMode
    'DrvFR.OperatorNumber
    'DrvFR.ECRSoftVersion
    'DrvFR.ECRBuild
    'DrvFR.ECRSoftDate
    'DrvFR.LogicalNumber
    'DrvFR.OpenDocumentNumber
    'DrvFR.ECRFlags
    'DrvFR.ReceiptRibbonIsPresent
    'DrvFR.JournalRibbonIsPresent
    'DrvFR.SlipDocumentIsPresent
    'DrvFR.SlipDocumentIsMoving
    'DrvFR.PointPosition
    'DrvFR.EKLZIsPresent
    'DrvFR.JournalRibbonOpticalSensor
    'DrvFR.ReceiptRibbonOpticalSensor
    'DrvFR.JournalRibbonLever
    'DrvFR.ReceiptRibbonLever
    'DrvFR.LidPositionSensor
    'DrvFR.IsPrinterLeftSensorFailure
    'DrvFR.IsPrinterRightSensorFailure
    'DrvFR.isDrawerOpen
    'DrvFR.ECRMode
    'DrvFR.ECRModeDescription
    'DrvFR.ECRMode8Status
    'DrvFR.ECRAdvancedMode
    'DrvFR.ECRAdvancedModeDescription
    'DrvFR.PortNumber
    'DrvFR.FMSoftVersion
    'DrvFR.FMBuild
    'DrvFR.FMSoftDate
    frmKKM.lbldateKKM.Caption = DrvFR.Date
    frmKKM.lblTimeKKM.Caption = DrvFR.Time
    'DrvFR.TimeStr
    'DrvFR.FMFlags
    'DrvFR.FM1IsPresent
    'DrvFR.FM2IsPresent
    'DrvFR.LicenseIsPresent
    'DrvFR.FMOverflow
    'DrvFR.BatteryCondition
    'DrvFR.SerialNumber
    'DrvFR.SessionNumber
    'DrvFR.FreeRecordInFM
    'DrvFR.RegistrationNumber
    'DrvFR.FreeRegistration
    'DrvFR.INN
    Exit Sub
err:
    ' MsgBox "������ � ��������� StatusKKM"

End Sub
'+++++++++ �����-��-� v.
Public Sub CheckKKM(�he�kType As Byte, GAS As Double, Cost As Double, Npost As Byte)
    '        On Error GoTo err
    '        '������� � ������ ������ ����� ���������� �� ���� ������
    '        With gGasStation(Npost)
    '            .gb�he�kType = �he�kType '��� ����������� ����:0-����� ������� ����;1,2-�� �����������;3-"������" ���; 4-��� ������
    '            .gdGasKKM = GAS 'CDbl(Format(Gas, "#####0.00")) '���������� ����, ������� ���������� �������� � ����
    '            .gdCostKKM = Cost '���� ��� ������ ����
    '            .gdFlagCheck = True '��������� ������� ������������� ������ ����
    '        End With
    '        Exit Sub
    'err:
    '        MsgBox "������ � ��������� CheckKKM"
    '        End

End Sub
Public Sub PrintCheckKKM(ByVal NumPost As Byte)    '��������� ��������� ��� � 3(���) ������� ��� �������� ����������� ��� � ������ �����

    '        Dim i As Integer
    '        Dim str As String
    '        Dim ModeKKM As Byte
    '        Dim ModeExKKM As Byte
    '        On Error GoTo err
    '        giErrorKKM = DrvFR.ResultCode
    '        gsErrorKKM = DrvFR.ResultCodeDescription
    '        StatusKKM
    '        If (giErrorKKM <> 0) And (giErrorKKM <> &H50) And (giErrorKKM <> &H8E) And (giErrorKKM <> &H1) And (giErrorKKM <> &H2) And (giErrorKKM <> &H6) Then '���� ������������ ������ �� �������� ���������:
    '        Else '���� ������ ���: ��� �������� ���������, �� �����:
    '            ModeKKM = DrvFR.ECRMode
    '            ModeExKKM = DrvFR.ECRAdvancedMode
    '            If ((ModeKKM = 2) Or (ModeKKM = 4)) And (ModeExKKM = 0) Then
    '                If gGasStation(NumPost).gdFlagCheck = True Then '���� ������ ���� �� ������,�� ���������
    '                    '���������� ���� �������
    '                    DrvFR.Password = pass
    '                    DrvFR.Quantity = gGasStation(NumPost).gdGasKKM ' * gGasStation(NumPost).gdCostKKM / gGasStation(NumPost).gdCostKKM
    '                    DrvFR.price = gGasStation(NumPost).gdCostKKM
    '                    DrvFR.Department = NumPost + 1
    '                    DrvFR.Tax1 = 1
    '                    DrvFR.Tax2 = 0
    '                    DrvFR.Tax3 = 0
    '                    DrvFR.Tax4 = 0
    '                    Select Case gGasStation(NumPost).gb�he�kType '��� ����������� ����
    '                    Case 0 '��� ����������� ����� ������������ ������� ���� ���������
    '                      DrvFR.StringForPrinting = "��� ���������"
    '                      DrvFR.Sale
    '                    Case 1 '��� �� �����������
    '                      Exit Sub
    '                    Case 2 '��� �� �����������
    '                      Exit Sub
    '                    Case 4 '!?! ��� ����������� ��� �������� ������������� ��������� �������� !?!
    '                    Case 3 '��� ���������� � ������ (��� ������� ����) � ������ ������ ���������: ��� ������� ���� ��������� ��� ������� ���� �� ��������� ��������
    '                    Case Else
    '                    End Select
    '                    '�������� ���� � �������(������� �� ��������� ������������ � ������� ������� ���)
    '                    DrvFR.Password = pass
    '                    DrvFR.CheckSubTotal '�������� ���� ����, �.�. � Summ1 ������� ��� Sale
    '                    DrvFR.Password = pass
    '                    If (gGasStation(NumPost).glTypeEnd = 2) And (gGasStation(NumPost).glTypeZapr = 0) And (gGasStation(NumPost).glSumm1 > 0) Then '������� �������� �� ����� �� �������
    '                        DrvFR.Summ1 = gGasStation(NumPost).glSumm1
    '                    End If
    '                    DrvFR.Summ2 = 0
    '                    DrvFR.Summ3 = 0
    '                    DrvFR.Summ4 = 0
    '                    DrvFR.DiscountOnCheck = 0 '������ ���
    '                    DrvFR.Tax1 = 1
    '                    DrvFR.Tax2 = 0
    '                    DrvFR.Tax3 = 0
    '                    DrvFR.Tax4 = 0
    '                    DrvFR.StringForPrinting = "===================================="
    '                    DrvFR.CloseCheck
    '                    Sleep 100
    '                    giErrorKKM = DrvFR.ResultCode
    '                    gsErrorKKM = DrvFR.ResultCodeDescription
    '                    '���� ��������� ������ �����������, �� ������� ���� ������ ����� ��������� ����
    '                    If (giErrorKKM = 0) Then ''And glOperator <> 0 Then
    '                        gGasStation(NumPost).gdFlagCheck = False
    '                        Set ZaprDN = StatDB.OpenRecordset("zapr", dbOpenDynaset) 'gas,number,price,operator,typeZapr,typeFinish,date=Max(Now) where number = " & Number)
    '                        ZaprDN.FindLast "number=" & NumPost & " and TYPEZAPR=0"
    '                        ZaprDN.Edit
    '                        ZaprDN("DATE") = Now
    '                        ZaprDN("TYPEZAPR") = 5 '������� ������� ����
    '                        ZaprDN.Update
    '                    Else
    '                    '!!!���� �� ���������������� ������ � ������� ��������������� ��������!!!
    '                    End If
    '
    '                    Exit Sub '� ����� ������ ������� �� ��������� ������, �.�. ������ �������� �������
    '                End If
    '            Else
    '            End If
    '        End If
    '        Exit Sub
    'err:
    '        MsgBox "������ � ��������� PrintCheckKKM"

End Sub
'+++++++++ �����-��-� v.03
Public Sub ConnectKKM()
    On Error GoTo err
    '������������ ������� ��, �������������� ��������� ��� � ����� c:\windiws\system
    '        Shell "regsvr32.exe /c /s " & Chr(34) '& "c:\windiws\drvfr.dll" & Chr(34)
    '��������� DrvFR ��� ����� ��������� � ��������
    Set DrvFR = CreateObject("AddIn.Drvfr")
    '� ������ ���������� �� ������ ������� ���������� DrvFR.��������
    DrvFR.GetActiveLD
    DrvFR.GetParamLD
    DrvFR.SetActiveLD
    StatusKKM
    If giErrorKKM <> 0 Then
    End If
    Exit Sub
err:
    MsgBox "��� �������� KKM"

End Sub



