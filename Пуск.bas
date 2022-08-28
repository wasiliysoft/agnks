Attribute VB_Name = "Module2"
Option Explicit
Public Function �������_ACL8113() As String

    Dim i           As Integer
    gla����� = Val("&H220")
    gl��������� = ISO813_DriverInit()
    If gl��������� <> ISO813_NoError Then
        i = MsgBox("Can not initial Driver!!!", , "ISO813 Card Error")
    ElseIf gl��������� = 2 Then
        �������_ACL8113 = "Driver open error !"
    Else

        �������_ACL8113 = "����� ACL8113 � �����"
    End If



End Function



Public Function �������_Pet48DIO() As String
    '�������������
    gl����� = Val("&H2C0")     '�������� �� ���������
    gl��������� = DIO_DriverInit(1)

    If gl��������� <> DIO_NoError Then
        MsgBox "Driver DIO Initialize OK!!"
    Else
        �������_Pet48DIO = "����� Pet48DIO � �����"
        ' Don't forget to close the driver by DIO_DriverClose()
    End If
    DIO_OutputByte gl����� + &H3, &H8B    '������������� CN1 : A0 -output, B0 & C0 - input
    DIO_OutputByte gl����� + &H7, &H8B    '������������� CN2 : A1 -output, B1 & C1 - input

    DIO_OutputByte gl�����, 0
    DIO_OutputByte gl����� + &H4, 0

    '��������� ���� 0 (���� A1)
    ' gl��������� = W_48DIO_DO(256, 0)

    '��������� ���� 0 (���� A0)
    ' gl��������� = W_48DIO_DO(0, 0)


End Function


Public Function �����������() As String
    ����������� = ""
    '������������� ����� ACL8113
    gs��������� = �������_ACL8113
    '������������� ����� Pet48DIO
    gs��������� = �������_Pet48DIO

    ����������� = "OK"
End Function



