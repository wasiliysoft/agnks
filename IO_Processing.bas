Attribute VB_Name = "IO_Processing"
Option Explicit


Public Sub ОпросПлат()
    If isDebug Then
        Обработка_1
        Обработка_1_debug
        Exit Sub
    End If
    'Прочитать состояние портов платы PET-48DIO
    update_gn48DIO

    'Опрос каналов платы ACL8113
    update_ggACL8113
End Sub


Public Sub Обработка_1()   
    Dim i           As Integer
    Dim Temp        As Double

    'Отобразить состояние датчиков, работающих с платой Pet48DIO
    update_gnДатчик

    'Пересчёт измеренных значений в ток:
    For i = 2 To 13
        gnDif(i) = gKi * (ggACL8113(i))
    Next i

    gnDif(15) = gKi * (ggACL8113(15))
    gnDif(14) = ggACL8113(14) * gKi_1

    gnDif(giChanel) = ggACL8113(giChanel) * ((2000 + 200) / 200)
    'Пересчёт измеренных значений в температуру (градусы Цельсия):

    For i = 8 To 13
        'Проверка (i - 4) , если не удовлетворяет то -1
        Temp = (gnDif(i) - 4)
        If Temp <= 17 And Temp >= -1 Then
            Select Case (i)
                Case 8
                    gnDif(i) = 200 * ((Temp + 1) / 18) - 50
                Case 9
                    gnDif(i) = 12.5 * Temp - 50
                Case 12
                    gnDif(i) = 6.25 * Temp - 50
                Case 13
                    gnDif(i) = 150 * ((Temp + 1) / 18) - 50
                Case Else
                    gnDif(i) = 6.25 * Temp - 50
            End Select
        Else
            gnDif(i) = -1
        End If
    Next i


    'Пересчёт измеренных значений в давление (в МПа):

    For i = 4 To 7
        'Проверка (i - 4) , если не удовлетворяет то -1
        Temp = (gnDif(i) - 4)
        If Temp <= 17 And Temp >= -1 Then
            gnDif(i) = gKp_1 * Temp
            If gnDif(i) < 0 Then
                gnDif(i) = 0
            End If
        Else
            gnDif(i) = -1
        End If
    Next i

    'Посмотреть для аккумуляторов нужен ли пересчет ?
    If (gnDif(15) - 4) <= 17 And (gnDif(15) - 4) >= -1 Then
        gnDif(15) = 0.1 / 16 * (gnDif(15) - 4)
    Else
        gnDif(15) = -1
    End If


    'Пересчет для ДД1.1 и ДД1.2
    For i = 2 To 3
        Temp = (gnDif(i) - 4)
        If Temp <= 17 And Temp >= -1 Then
            gnDif(i) = Temp * gKp
        Else
            gnDif(i) = -1
        End If
    Next i

    ' Расчет оборотов ДВС
    If (gnDif(14) - 4) <= 17 And (gnDif(14) - 4) >= -1 Then
        gnDif(14) = (gnDif(14) - 4) * 200   ' =3200/16
    Else
        gnDif(14) = -1
    End If
    Call AddSensorsData(2, gnDif(5), gnDif(11), gnDif(4), 1.5, 0.95 * gdK, 0)
    gdИР2 = GetMass(2)
    Temp = GetMassExpense(2)
    'Считать расход (общий) по ИР1
    If giMainРасход = 1 Then
        Call AddSensorsData(1, gnDif(2), gnDif(9), gnDif(3), 6, 0.95 * gdK, 0)
    Else
        Call AddSensorsData(1, gnDif(2), gnDif(9), gnDif(3), 6, 0.95 * gdK, -(Temp))
    End If
    gdИР1 = GetMass(1)
End Sub


Private Sub Обработка_1_debug()
    Dim i As Integer

    gnDif(4) = 10
   'Debug.Print "giStage2", giStage2
    If (giStage2 = 0) Then
         gnDif(5) = 10
    End If
    
    If giStage2 = 9 Then
        gnDif(4) = gnDif(5) - 0.5
        gnDif(5) = gnDif(5) + 0.3
        If (gnDif(5) > 20) Then
            gnDif(4) = gnDif(5)
        End If
        gdИР2 = gdИР2 + 0.15
        
       ' Debug.Print "gnDif(4)", gnDif(4)
       ' Debug.Print "gnDif(5)", gnDif(5)
        
    End If
     '0 A0 Output 0-7
'    gnДатчик(0).Data = 0        ' Управление Реле 1 (контроль)
    'gnДатчик(1).Data = 0        ' Управление К7
     '1 B0 Input 8-15
'    gnДатчик(8).Data = 0        ' Вход 1 (контроль)
'    gnДатчик(9).Data = 0        ' Приборметанавтехотсек
'    gnДатчик(10).Data = 0
'    gnДатчик(11).Data = 0
'    gnДатчик(12).Data = 0
'    gnДатчик(13).Data = 0
'    gnДатчик(14).Data = 0
'    gnДатчик(15).Data = 1          ' Автомат. упр-е
     '2 C0 Input 16-23
'    gnДатчик(16).Data = 0       ' К2 открыт   0
'    gnДатчик(17).Data = 0       ' К3 открыт   1
'    gnДатчик(18).Data = 0       ' К4 открыт   2
'    gnДатчик(19).Data = 0       ' К5 открыт   3
'    gnДатчик(20).Data = 0       ' К6 открыт   4
'    gnДатчик(21).Data = 0       ' К1 открыт   5
'    gnДатчик(22).Data = 0       '             6
'    gnДатчик(23).Data = 0       ' К7 открыт   7
    '3 Config address
    '4 A1 Output 24-31
'    gnДатчик(24).Data = 0       ' Управление Реле 2 (контроль)
'    gnДатчик(25).Data = 0       ' Стоп двигАГНКС
'    gnДатчик(26).Data = 0       ' Открыть КЭ1
'    gnДатчик(27).Data = 0       ' Открыть КЭ2
'    gnДатчик(28).Data = 0       ' Открыть КЭ3
'    gnДатчик(29).Data = 0       ' Открыть КЭ4
'    gnДатчик(30).Data = 0       ' Открыть КЭ5
'    gnДатчик(31).Data = 0       ' Открыть КЭ6
     '5 B1 Input 32-39
'    gnДатчик(32).Data = 0       ' Вход 2(контроль)
'    gnДатчик(33).Data = 0       ' Высокая tC ОЖид. ДВС
'    gnДатчик(34).Data = 0       ' Разряд аккумулятора
'    gnДатчик(35).Data = 0       ' Pmax масла ДВС
'    gnДатчик(36).Data = 0       ' Муфта сцепления
'    gnДатчик(37).Data = 0       ' Охл. ДВС(вентилятор)
'    gnДатчик(38).Data = 0       ' Авар.вытяж.вентил.В1
'    gnДатчик(39).Data = 0       ' Авар.вытяж.вентил.В2
     '6 C1 Input 40-47
'    gnДатчик(40).Data = 0       ' Потеря напряжения
'    gnДатчик(41).Data = 0       ' Метан 10% (Отсек ДВС)
'    gnДатчик(42).Data = 0       ' Метан 20% (Отсек ДВС)
'    gnДатчик(43).Data = 0       ' Метан 10% (Техн.отсек)
'    gnДатчик(44).Data = 0       ' Метан 20% НКПР
'    gnДатчик(45).Data = 0       ' Пожар в отсеке ДВС
'    gnДатчик(46).Data = 0       ' Пожар в тех.отсеке
'    gnДатчик(47).Data = 0       ' Отказ СТМ-10

    gnDif(2) = 1000 ' ДД1.1
    gnDif(3) = 1000 ' ДД1.2
'    gnDif(4) = 0 ' ДД2.1
'    gnDif(5) = 0 ' ДД2.2
    gnDif(6) = 22 ' ДД6, компрессор
    gnDif(7) = 20 ' ДД8, аккомулятор
    gnDif(8) = 10 ' ДТ1, датчик температуры
    gnDif(9) = 10 ' ДТ1.1, датчик температуры
    gnDif(10) = 10 ' ДТ2, датчик температуры
    gnDif(11) = 10 ' ДТ2.1, датчик температуры
    gnDif(12) = 10 ' ДТ3, датчик температуры
    gnDif(13) = 10 ' ДТ4, датчик температуры на выходе компрессора
    'gnDif(14) = 0 ' Обороты ДВС
    gnDif(15) = 230 ' ДД4
    gnDif(16) = 24.4 ' Напряжение АКБ
   ' Debug.Print gnDif(0)
End Sub


