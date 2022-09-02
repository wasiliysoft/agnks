Attribute VB_Name = "IO_Digital"
Option Explicit

Private Const glАдрес = &H2C0
Private Const configCN1 = glАдрес + &H3  '   707
Private Const configCN2 = glАдрес + &H7  '   711

Global Const A0 = glАдрес + &H0     'Порты 704
Global Const B0 = glАдрес + &H1     '705
Global Const C0 = glАдрес + &H2     '706

Global Const A1 = glАдрес + &H4     'Порты 708
Global Const B1 = glАдрес + &H5     '709
Global Const C1 = glАдрес + &H6     '710

' The Driver functions
Private Declare Function DIO_DriverInit Lib "DIO.DLL" (wTotalBoards As Integer) As Integer
Declare Sub DIO_DriverClose Lib "DIO.DLL" ()

' The DIO functions
Private Declare Sub DIO_OutputByte Lib "DIO.DLL" _
        (ByVal address As Integer, ByVal dataout As Byte)
Private Declare Function DIO_InputByte Lib "DIO.DLL" _
        (ByVal address As Integer) As Integer


Public gnДатчик(48) As Sensor    'состояние датчиков по платам TB-24P и TB-16P8R

Private gn48DIO(5)   As Long    'состояние регистров платы PET-48DIO


Public Function Init_DIO_Driver() As String
    Dim i As Integer
    Dim msg As String
    i = DIO_DriverInit(1)
    Select Case i
        Case 0: msg = "NoError"
        Case 1: msg = "DriverOpenError"
        Case 2: msg = "DriverNoOpen"
        Case 3: msg = "GetDriverVersionError"
        Case 4: msg = "InstallIrqError"
        Case 5: msg = "ClearIntCountError"
        Case 6: msg = "GetIntCountError"
        Case 7: msg = "ResetError"
        Case 8: msg = "RemoveIrqError"
        Case 9: msg = "GetTotalBoardError"
        Case 10: msg = "CardNotFound"
        Case 11: msg = "GetConfigError"
        Case 12: msg = "ExceedBoardNumber"
    End Select

    If i <> 0 Then
        MsgBox msg, vbExclamation, "Driver DIO"
    End If

    DIO_OutputByte configCN1, &H8B    'Устанавливаем CN1 : A0 -output, B0 & C0 - input
    DIO_OutputByte configCN2, &H8B    'Устанавливаем CN2 : A1 -output, B1 & C1 - input
    
    ' Выключить реле
    ' TODO можно переделать на ROff
    DIO_OutputByte A0, 0
    DIO_OutputByte A1, 0
End Function

Public Sub update_gn48DIO()
    gn48DIO(0) = CInt(DIO_InputByte(A0))
    gn48DIO(1) = Not (CInt(DIO_InputByte(B0)))
    gn48DIO(2) = Not (CInt(DIO_InputByte(C0)))

    gn48DIO(3) = CInt(DIO_InputByte(A1))
    gn48DIO(4) = Not (CInt(DIO_InputByte(B1)))
    gn48DIO(5) = Not (CInt(DIO_InputByte(C1)))
End Sub

Public Sub update_gnДатчик()
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

' флаг ручного управления, true если ручное управление
Function isHandControl() As Boolean
    isHandControl = Not (CBool(gnДатчик(15).Data))
End Function

' К1 открыт? true если открыт
Function k1_isOpen() As Boolean
    k1_isOpen = CBool(gnДатчик(21).Data)
End Function

' К2 открыт? true если открыт
Function k2_isOpen() As Boolean
    k2_isOpen = CBool(gnДатчик(16).Data)
End Function

' К3 открыт? true если открыт
Function k3_isOpen() As Boolean
    k3_isOpen = CBool(gnДатчик(17).Data)
End Function

' К4 открыт? true если открыт
Function k4_isOpen() As Boolean
    k4_isOpen = CBool(gnДатчик(18).Data)
End Function

' К5 открыт? true если открыт
Function k5_isOpen() As Boolean
    k5_isOpen = CBool(gnДатчик(19).Data)
End Function

' К6 открыт? true если открыт
Function k6_isOpen() As Boolean
    k6_isOpen = CBool(gnДатчик(20).Data)
End Function

' К7 открыт? true если открыт
Function k7_isOpen() As Boolean
    k7_isOpen = CBool(gnДатчик(23).Data)
End Function

' Муфта сцепления вкл?
Function isClutchOn() As Boolean
    isClutchOn = CBool(gnДатчик(36).Data)
End Function
'Функция выводит в port 1
Public Sub ROn(port As Integer, n As Integer)
    Dim b As Byte
    ' текущее состояние
    b = getSoftPortState(port)
    ' Битовое ИЛИ (1 останутся только если они есть в обоих байтах)
    ' Битовое ИЛИ (1 останутся из обоих байтов)
    b = b Or n

    If (isDebug) Then
        Debug.Print "Запись 1 в адрес: " & port & " n: " & n
        gn48DIO(getIndexByPort(port)) = b
        Ron_debug port, n
    Else
        DIO_OutputByte port, b
    End If
    ' Выполнить опрос платы DIO
    ОпросПлат
End Sub

'Функция выводит в port 0
Public Sub ROff(port As Integer, n As Integer)
    Dim b As Byte
    ' текущее состояние
    b = getSoftPortState(port)
    ' Битовое И (1 останутся только если они есть в обоих байтах)
    b = b And n

    If (isDebug) Then
        Debug.Print "Запись 0 в адрес: " & port & " n: " & n
        gn48DIO(getIndexByPort(port)) = b
        Roff_debug port, n
    Else
        DIO_OutputByte port, b
    End If
    ' Выполнить опрос платы DIO
    ОпросПлат
End Sub

' Возвращает состояние порта на момент последнего опроса
Private Function getSoftPortState(port As Integer) As Byte
    Dim i As Integer: i = getIndexByPort(port)
    getSoftPortState = gn48DIO(i)
    'If isDebug Then
    '    Debug.Print "getSoftPortState", "port " & port, "return " & getSoftPortState
    'End If
End Function

Private Function getIndexByPort(port As Integer) As Integer
    Select Case port
        Case A0: getIndexByPort = 0
        Case B0: getIndexByPort = 1
        Case C0: getIndexByPort = 2
        Case A1: getIndexByPort = 3
        Case B1: getIndexByPort = 4
        Case C1: getIndexByPort = 5
        Case Else: err.Raise -1, , "Некорректный адрес порта: " & port
    End Select
End Function


Private Sub Ron_debug(port As Integer, n As Integer)
    If (port = A1 And n = 2) Then ' Стоп ДВС
        ggACL8113(14) = 0.8
        Debug.Print "ДВС остановлен"
    ElseIf (port = A1 And n = 4) Then ' Открыть К1
        gn48DIO(getIndexByPort(C0)) = getSoftPortState(C0) Or 32
        Debug.Print "КЭ1 открыт"
    ElseIf (port = A1 And n = 6) Then ' Открыть К1 и стоп ДВС
        ggACL8113(14) = 0.8
        Debug.Print "ДВС остановлен"
        gn48DIO(getIndexByPort(C0)) = getSoftPortState(C0) Or 32
        Debug.Print "КЭ1 открыт"
    ElseIf (port = A1 And n = 8) Then
        gn48DIO(getIndexByPort(C0)) = getSoftPortState(C0) Or 1
        Debug.Print "КЭ2 открыт"
    ElseIf (port = A1 And n = 16) Then
        gn48DIO(getIndexByPort(C0)) = getSoftPortState(C0) Or 2
        Debug.Print "КЭ3 открыт"
    ElseIf (port = A1 And n = 24) Then ' Открыть К2 и К3
        gn48DIO(getIndexByPort(C0)) = getSoftPortState(C0) Or 1
        Debug.Print "КЭ2 открыт"
        gn48DIO(getIndexByPort(C0)) = getSoftPortState(C0) Or 2
        Debug.Print "КЭ3 открыт"
    ElseIf (port = A1 And n = 32) Then
        gn48DIO(getIndexByPort(C0)) = getSoftPortState(C0) Or 4
        Debug.Print "КЭ4 открыт"
    ElseIf (port = A1 And n = 34) Then 'Стоп ДВС, открыть КЭМ4
        ggACL8113(14) = 0.8
        Debug.Print "ДВС остановлен"
        gn48DIO(getIndexByPort(C0)) = getSoftPortState(C0) Or 4
        Debug.Print "КЭ4 открыт"
    ElseIf (port = A1 And n = 42) Then ' Cтоп ДВС, открыть КЭМ2, открыть КЭМ4
        ggACL8113(14) = 0.8
        Debug.Print "ДВС остановлен"
        gn48DIO(getIndexByPort(C0)) = getSoftPortState(C0) Or 1
        Debug.Print "КЭ2 открыт"
        gn48DIO(getIndexByPort(C0)) = getSoftPortState(C0) Or 4
        Debug.Print "КЭ4 открыт"
    ElseIf (port = A1 And n = 64) Then
        gn48DIO(getIndexByPort(C0)) = getSoftPortState(C0) Or 8
        Debug.Print "КЭ5 открыт"
    ElseIf (port = A1 And n = 128) Then
        gn48DIO(getIndexByPort(C0)) = getSoftPortState(C0) Or 16
        Debug.Print "КЭ6 открыт"
    ElseIf (port = A0 And n = 2) Then
        gn48DIO(getIndexByPort(C0)) = getSoftPortState(C0) Or 128
        Debug.Print "КЭ7 открыт"
    ElseIf (port = B0 And n = 128) Then
        Debug.Print "Выкл. ручное упарвление"
    ElseIf (port = B1 And n = 16) Then
        Debug.Print "Вкл муфта сцепления"
    Else
        Debug.Print "Необработанная команда Открыть", port, n
    End If
End Sub

Private Sub Roff_debug(port As Integer, n As Integer)
    If (port = A1 And n = 0) Then
        gn48DIO(getIndexByPort(C0)) = 0
        Debug.Print "Все КМ закрыты"
    ElseIf (port = A1 And n = 1) Then ' Закрыть все КМ, вкл Реле 2
        gn48DIO(getIndexByPort(C0)) = 0
        Debug.Print "Все КМ закрыты, Реле 2 ВКЛ"
    ElseIf (port = A1 And n = 239) Then ' Закрыть К3
        gn48DIO(getIndexByPort(C0)) = getSoftPortState(C0) And Not (2)
        Debug.Print "К3 Закрыт"
    ElseIf (port = A1 And n = 223) Then ' Закрыть К4
        gn48DIO(getIndexByPort(C0)) = getSoftPortState(C0) And Not (4)
        Debug.Print "К4 Закрыт"
    ElseIf (port = A1 And n = 127) Then ' Закрыть К6
        gn48DIO(getIndexByPort(C0)) = getSoftPortState(C0) And Not (16)
        Debug.Print "К6 Закрыт"
    ElseIf (port = A1 And n = 191) Then ' Закрыть К5
        gn48DIO(getIndexByPort(C0)) = getSoftPortState(C0) And Not (8)
        Debug.Print "К5 Закрыт"
    ElseIf (port = B1 And n = Not (16)) Then
        Debug.Print "Выкл муфта сцепления"
    Else
        Debug.Print "Необработанная команда Закрыть", port, n
    End If
End Sub



' TODO Проверить корректность
' ROff A1, 1 Не только закрывает К1-6, но и ВЫКЛЮЧАЕТ реле "СТОП ДВИГ"