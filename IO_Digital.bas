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
    'Инициализация
    glАдрес = Val("&H2C0")     'Оставляю по умолчанию
    glРезультат = DIO_DriverInit(1)

    If glРезультат <> DIO_NoError Then
        MsgBox "Driver DIO Initialize OK!!"
    Else
        Init_DIO_Driver = "Плата Pet48DIO в норме"
        ' Don't forget to close the driver by DIO_DriverClose()
    End If
    DIO_OutputByte glАдрес + &H3, &H8B    'Устанавливаем CN1 : A0 -output, B0 & C0 - input
    DIO_OutputByte glАдрес + &H7, &H8B    'Устанавливаем CN2 : A1 -output, B1 & C1 - input

    DIO_OutputByte glАдрес, 0
    DIO_OutputByte glАдрес + &H4, 0

    'Выключить реле 0 (порт A1)
    ' glРезультат = W_48DIO_DO(256, 0)

    'Выключить реле 0 (порт A0)
    ' glРезультат = W_48DIO_DO(0, 0)


End Function
