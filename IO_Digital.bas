Attribute VB_Name = "IO_Digital"
Option Explicit
Global Const DIO_NoError = 0
Global Const DIO_DriverOpenError = 1
Global Const DIO_DriverNoOpen = 2
Global Const DIO_GetDriverVersionError = 3
Global Const DIO_InstallIrqError = 4
Global Const DIO_ClearIntCountError = 5
Global Const DIO_GetIntCountError = 6
Global Const DIO_ResetError = 7
Global Const DIO_RemoveIrqError = 8

Global Const DIO_GetTotalBoardError = 9
Global Const DIO_CardNotFound = 10
Global Const DIO_GetConfigError = 11
Global Const DIO_ExceedBoardNumber = 12
 

' The Driver functions
Declare Function DIO_DriverInit Lib "DIO.DLL" _
        (wTotalBoards As Integer) As Integer
Declare Sub DIO_DriverClose Lib "DIO.DLL" ()

' The DIO functions
Declare Sub DIO_OutputByte Lib "DIO.DLL" _
        (ByVal address As Integer, ByVal dataout As Byte)
Declare Function DIO_InputByte Lib "DIO.DLL" _
        (ByVal address As Integer) As Integer




Public Function Init_DIO_Driver() As String
    '�������������
    gl����� = Val("&H2C0")     '�������� �� ���������
    gl��������� = DIO_DriverInit(1)

    If gl��������� <> DIO_NoError Then
        MsgBox "Driver DIO Initialize OK!!"
    Else
        Init_DIO_Driver = "����� Pet48DIO � �����"
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
