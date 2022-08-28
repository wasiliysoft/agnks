Attribute VB_Name = "IO_Digital"
Option Explicit

Private Const glАдрес = &H2C0

Private Const DIO_NoError = 0
Private Const DIO_DriverOpenError = 1
Private Const DIO_DriverNoOpen = 2
Private Const DIO_GetDriverVersionError = 3
Private Const DIO_InstallIrqError = 4
Private Const DIO_ClearIntCountError = 5
Private Const DIO_GetIntCountError = 6
Private Const DIO_ResetError = 7
Private Const DIO_RemoveIrqError = 8

Private Const DIO_GetTotalBoardError = 9
Private Const DIO_CardNotFound = 10
Private Const DIO_GetConfigError = 11
Private Const DIO_ExceedBoardNumber = 12


' The Driver functions
Declare Function DIO_DriverInit Lib "DIO.DLL" _
        (wTotalBoards As Integer) As Integer
Declare Sub DIO_DriverClose Lib "DIO.DLL" ()

' The DIO functions
Declare Sub DIO_OutputByte Lib "DIO.DLL" _
        (ByVal address As Integer, ByVal dataout As Byte)
Declare Function DIO_InputByte Lib "DIO.DLL" _
        (ByVal address As Integer) As Integer


Public gn48DIO(5)   As Long    'состояние регистров платы PET-48DIO
Public gnДатчик(48) As Sensor    'состояние датчиков по платам TB-24P и TB-16P8R

Public Function Init_DIO_Driver() As String
    'Инициализация
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

Sub update_gn48DIO()
    gn48DIO(0) = CInt(DIO_InputByte(glАдрес))
    gn48DIO(1) = Not (CInt(DIO_InputByte(glАдрес + 1)))
    gn48DIO(2) = Not (CInt(DIO_InputByte(glАдрес + 2)))

    gn48DIO(3) = CInt(DIO_InputByte(glАдрес + 4))
    gn48DIO(4) = Not (CInt(DIO_InputByte(glАдрес + 5)))
    gn48DIO(5) = Not (CInt(DIO_InputByte(glАдрес + 6)))
End Sub

Sub update_gnДатчик()
    Dim p           As Integer
    Dim r           As Integer
    Dim i           As Integer
    Dim j           As Integer
    For i = 0 To 5
        p = gn48DIO(i)
        For j = 0 To 7
            r = p Mod 2
            If r = 0 Then
                gnДатчик(8 * i + j).Data = 0
            Else
                gnДатчик(8 * i + j).Data = 1
            End If
            p = Int(p / 2)
        Next j
    Next i
End Sub

'Функция выводит в port 0
Public Function ROff(port As Integer, n As Integer) As Integer
    Dim t           As Byte
    Dim j           As Integer

    Select Case port
        Case A0
            j = 0
        Case B0
            j = 1
        Case C0
            j = 2
        Case A1
            j = 4
        Case B1
            j = 5
        Case C1
            j = 6
        Case Else    '''Возможно и не надо !!!
            j = 3
    End Select

    If j <= 2 Then
        t = gn48DIO(j)    'считываем состояние порта
    Else
        t = gn48DIO(j - 1)    'считываем состояние порта
    End If
    'Закрыть
    t = t And n  ' 0 в n-ый канал
    ''''Для отладки !!!
    'W_48DIO_DO port, t
    DIO_OutputByte glАдрес + j, t

    gn48DIO(j) = t
    ОпросПлат  'Нужно чтобы узнать ИР2
    '----------
    '----------
End Function

'Функция выводит в port 1
Public Function ROn(port As Integer, n As Integer) As Integer
    Dim t           As Byte
    Dim j           As Integer

    Select Case port
        Case A0
            j = 0
        Case B0
            j = 1
        Case C0
            j = 2
        Case A1
            j = 4
        Case B1
            j = 5
        Case C1
            j = 6
        Case Else    '''Возможно и не надо !!!
            j = 3
    End Select
    If j <= 2 Then
        t = gn48DIO(j)    'считываем состояние порта
    Else
        t = gn48DIO(j - 1)    'считываем состояние порта
    End If
    'Открыть
    t = t Or n  ' 1 в n-ый канал
    ''''Для отладки !!!
    '     W_48DIO_DO port, t
    't = 2
    DIO_OutputByte glАдрес + j, t

    gn48DIO(j) = t
    '----------
    ОпросПлат
End Function