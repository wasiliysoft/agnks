Attribute VB_Name = "Main"
Option Explicit


'Функция обработки аварийных ситуаций
Public Function Danger() As String
    ' TODO АвтоАктивация вкладки СХЕМА
    ' FIXME выполняется проверка управляющей команды!!!
    If gnДатчик(29).Data = 1 Then
        'Если перепад во вх. и вых. рукове станет меньше 5 кг
        If Abs(gnDif(6) - gnDif(2)) <= 0.5 Then
            ROff A1, 223    'Закрыть К4
            If isFireTech Then
                ROn A1, 24 '(16 + 8) Открыть КЭМ3 и КЭМ2
                ROn A0, 2 'Открыть КЭМ7
            Else
                ROn A1, 16 'Открыть КЭМ3
            End If
        End If
    End If
    frmStart.cmdDanger.Visible = True
End Function



'Функция остановки АГНКС
Public Function ОстановАГНКС() As String
    Dim s           As String
    ROff A1, 1 'Закрыть К 1-6
    'если загазованность 20 %
    If (gnДатчик(42).Data = 1) Or (gnДатчик(44).Data = 1) Then
        ROn A1, 42 '(2+8+32) Стоп ДВС, открыть К2, открыть К4
        ROn A0, 2  'Открыть К7
    ElseIf isFireTech Then    'если пожар в техническом отсеке
        ROn A1, 34 '(2+32) Стоп ДВС, открыть КЭМ4
    End If


    giStage2 = 0
    giStage = 3  'Переход на этап Danger
    giStage1 = 0
    gbAkkum = False
    frmStart.SSCmdStart.Enabled = False
    gbCmdStart = True
    StopOutput (2)
    gbStopAGNKS = True
    ОстановАГНКС = "Останов АГНКС"
End Function

'Функция остановки ДВС
Public Function ОстановДВС() As String
    'TODO Зачем закрывать пистолет при остановке ДВС?
    ROff A1, 191 'Закрыть К5 (пистолет)

    ROn A1, 32 'Открыть К4

    giStage2 = 0
    giStage = 1  'Переход на этап ИсхСост
    giStage1 = 1
    giMainРасход = 0

    gbAkkum = False
    frmStart.SSCmdStart.Enabled = False
    gbCmdStart = False
End Function
'Сама процедура заправки
Public Function Заправка()
    Dim s, s1       As String
    Dim MaxIR       As Double    'Запоминаем max расход при открытии КЭМ6
    Dim p           As Double

    ' ПОДЭТАП 8  - Заправка только от ов
    If giStage2 = 8 Then
        'Заправка машин

        If k4_isOpen Then
            ROff A1, 223 'Закрыть К4
        Else
            ROff A1, 239 'Закрыть К3
        End If


        ROn A1, 64 'Открыть КЭ5
        giStage2 = 9
        ROn A1, 128 'Открыть КЭ6
        gdРасход1 = 0    'Обнуляем расход на одну машину
        ResetExpenseCounter_2
        StartOutput (2)
        gbDontStat = True    'Нельзя работать с диском
        Exit Function
    End If

    'ПодЭтап 9
    If giStage2 = 9 Then
        If (Abs(gnDif(5) - gnDif(4)) > 0.5) Then
            Заправка = "Идет заправка "
            'Считаем расход на одну машину (за полсекунды)
            gdTime = GetTimeCounter_2
            gdРасход1 = gdИР2
            Exit Function
        Else
            ROff A1, 191 'Закрыть К5 (пмстолет)
            ROff A1, 127 'Закрыть К4 (Акк)
            StopOutput (2)
            gbDontStat = False    'Можно работать с диском

            gdTime = GetTimeCounter_2

            'Заполнить статистику по заправке

            '<<<<Прекратить считать расход>>>>
            StatRS_Insert

            s = Format(Now, "hh:mm:ss") + "        " + Format((gdРасход1 / gdPlot), "###0.00")
            frmStart.lstStat(0).AddItem s

            gbЗаправка = False

            'Разрешить повторную заправку автомобиля во время заправки аккумуляторов
            frmStart.SSCmdStart.Enabled = True
            gbAkkum = True
            giStage = 1    'Переход на Этап Предпуска
            giStage1 = 0
            giStage2 = 0
            Exit Function

        End If
    End If


    ' ПОДЭТАП 1
    If (giStage2 = 0) And (gbFrmShow = False) Then
        frmЗапрос.Show vbModeless
        gbFrmShow = True
        Заправка = "Выведен диалог (пистолет)"
        Exit Function
    End If

    ' ПОДЭТАП 2
    If (giStage2 = 1) And (gbFrmShow = False) Then
        If giTrigger = 0 Then ' Пистолет не вставлен
            giStage2 = 0
            gbЗаправка = True
            gbAkkum = False
            giStage = 1    'Переход на этап ПредПуск
            giStage1 = 1    'Сразу на проверку ДВС и компрессора

            Заправка = "Переход на этап ПредПуск"
            Exit Function
        Else ' Пистолет вставлен
            giStage2 = 2
            Car = 1
            s = "Заправка машин"
        End If
    End If

    ' ПОДЭТАП 3
    If giStage2 = 2 Then
        If Car = 1 Then
            'Заправка машин
            ROn A1, 64  'Открыть КЭ5
            giStage2 = 3
            gdРасход1 = 0    'Обнуляем расход на одну машину
            ResetExpenseCounter_2
            StartOutput (2)
            gbDontStat = True    'Нельзя работать с диском
            giMainРасход = 1
        End If
    End If

    ' ПОДЭТАП 4
    If giStage2 = 3 Then
        'Считать расход заправки автомобиля
        gbЗаправка = True

        If (gnDif(7) - gnDif(5)) >= 2 Then    'Разница давлений в аккумуляторах и баке
            ROn A1, 128 'Открыть К6 - заправка и от аккумуляторов
        End If

        If k4_isOpen Then
            ROff A1, 223 'Закрыть К4 - Начинает гнать газ компрессор
        Else
            ROff A1, 239 'Закрыть К3 - Начинает гнать газ компрессор
        End If

        giStage2 = 4    'Переходим к подэтапу заправки аккумуляторов
        ОпросПлат
        Обработка_1
    End If

    ' ПОДЭТАП 5
    If giStage2 = 4 Then
        '<<<<Считать расход по ИР2>>>>


        MaxIR = GetMassExpense_2
        If (gbAkkum = False) And ((k6_isOpen And (((MaxIR * 3600) <= gdRashAkkEnd) _
                And (MaxIR > 0)) And (MaxIR >= 5)) Or ((gnDif(7) - gnDif(4)) <= 0.5)) Then           
            ROff A1, 127 'Закрыть К6 (Акк)
            'Exit Function
        End If
        If (gbAkkum = False) And ((Not (gnDif(4) >= gdUpLevel))) Then
            Заправка = "Идет заправка "
            'Считаем расход на одну машину (за полсекунды)
            gdРасход1 = gdИР2
            gdTime = GetTimeCounter_2
            Exit Function
        ElseIf (gbAkkum = False) Then
            ROff A1, 191 'Закрыть К5 (пистолет)
            gbDontStat = False    'Можно работать с диском
            StopOutput (2)
            gdTime = GetTimeCounter_2
            'Заполнить статистику по заправке
            StatRS_Insert

            s = Format(Now, "hh:mm:ss") + "        " + Format((gdРасход1 / gdPlot), "###0.00")
            frmStart.lstStat(0).AddItem s

            '<<<<Прекратить считать расход>>>>
            gbЗаправка = False            
            ROn A1, 128 'Открыть КЭ6 ЗАПРАВЛЯЕМ АККУМУЛЯТОРЫ

            'Разрешить повторную заправку автомобиля во время заправки аккумуляторов
            frmStart.SSCmdStart.Enabled = True
            gbAkkum = True
        End If

        'заправлять аккумуляторы до 200 кг
        If (gnDif(7) < gdUpLevel) And (gbAkkum = True) Then
            Заправка = "Заправка аккумуляторов"
            Exit Function
        Else
            ROff A1, 127 'Закрыть К6 (Акк)
        End If
        'Если выведена форма запроса о заправке машины
        If gbFrmShow = True Then
            frmЗапрос.Hide
            frmStart.SSCmdStart.Enabled = True
            gbFrmShow = False
        End If

        'Выключить Двигатель
        s = ОстановДВС
        '<<<<Прекратить считать расход>>>>
        gbЗаправка = False


    End If


    ' ПОДЭТАП 7  - во время заправки аккумуляторов переход на заправку машин
    If giStage2 = 7 Then       
        ROn A1, 64 'Открыть КЭ5
        s = "Переходим на заправку машин"
        ResetExpenseCounter_2
        StartOutput (2)
        ОпросПлат
        Обработка_1

        giStage2 = 4
        'Считать расход заправки автомобиля
        giMainРасход = 1
        gbЗаправка = True
        gbAkkum = False
        gdРасход1 = 0    'Обнуляем расход на одну машину

        gdTime = GetTimeCounter_2
    End If


    Заправка = s
End Function


'Приводит АГНКС в исходное состояние
Public Function ИсхСост() As String
    Dim s           As String
    Dim norma       As Boolean
    frmStart.SSCommand2(1).Enabled = True
    frmStart.SSCommand2(0).Enabled = True

    s = ""
    norma = True
    gbRunDVS = False
    ' TODO проверить утверждение ниже, пропущен k7
    'Входные реле включены (порт A0 и A1) ? - неисправны
    If k2_isOpen Or k3_isOpen Or k4_isOpen Or _
            k5_isOpen Or k6_isOpen Or k1_isOpen Then
        s = "Есть открытые КЭМы !!!"
        norma = False
    End If

    ' FIXME выполняется проверка управляющей команды!!!
    ' но это тут кажется это не испарвить
    If (gnДатчик(25).Data = 1) Then
        s = s & "Нажата Останов ДВС !!!"
        norma = False
    End If

    ' If gnDif(1) <= 0.3 Then   '0.3 потому что плавает показание
    '   s = s & "Нет газа !!!"
    '   norma = False
    ' End If

    If isClutchOn Then
        s = s & "Включена муфта !!!"
        norma = False
    End If

    If gnDif(14) > 100 Then  'Есть обороты
        s = s & "Есть обороты ДВС !!!"
        norma = False
    End If

    If norma Then
        s = "АГНКС в исходном состоянии ."
        frmStart.SSCmdStart.Enabled = True
        gbOnlyAkk = True
    Else
        s = s & "АГНКС не готова !!!"
        frmStart.SSCmdStart.Enabled = False
    End If
    ИсхСост = s
End Function


'Подготавливает компрессор к пуску
Public Function ПредПуск() As String
    If giStage1 = 0 Then
        'Если есть давление в выходном трубопроводе , то открыть КЭМ4
        If (gnDif(6) - gnDif(2)) >= 0.25 Then
            ROn A1, 32 'Открыть К4
        Else
            ROn A1, 16 'Открыть К3 - для запуска ДВС
        End If
        gbAkkum = True
        giStage1 = 1    ' Пререход на второй подэтап
    End If

    If giStage1 = 1 Then
        'ВТОРОЙ ПОДЭТАП
        If gnDif(14) < 100 Then
            'Нет оборотов ДВС
            gbAkkum = True
            frmStart.SSCmdStart.Enabled = True
            'Если ДВС был запущен и заглох , то переход на Этап ИсхСост
            If gbRunDVS = True Then
                giStage2 = 0
                giStage = 0    'Переход на этап ИсхСост
                giStage1 = 0
                gbAkkum = False
                gbRunDVS = False
                frmStart.SSCmdStart.Enabled = False
                gbCmdStart = True

                'frmStart.Timer2.Enabled = False
                'TODO проверить коррекность gnДатчик(25)
                If k2_isOpen Or k3_isOpen Or k4_isOpen Or _
                        k5_isOpen Or k6_isOpen Or k1_isOpen Or _
                        (gnДатчик(25).Data = 1) Then
                    ROff A1, 1 'закрыть К 1-6
                End If

            End If

            If isClutchOn Then
                ПредПуск = "Двигатель не готов к запуску !!! Включена муфта "
                Exit Function
            Else
                ПредПуск = "Двигатель готов к запуску !!!"
                Exit Function
            End If
            'Есть обороты ДВС
        ElseIf Not (isClutchOn) Then
            ПредПуск = "Двигатель на холостом ходу !!!"
            frmStart.SSCmdStart.Enabled = False
            gbOnlyAkk = False
            gbAkkum = False
            Exit Function
        ElseIf gnDif(14) <= 1700 Then
            ПредПуск = "ДВС не вышел на рабочий режим !!!"
            frmStart.SSCmdStart.Enabled = False
            gbRunDVS = True
            gbOnlyAkk = False
            gbAkkum = False
            Exit Function
        Else
            ПредПуск = "Компрессор в работе, можно заправлять !!!"
            frmStart.SSCmdStart.Enabled = True
            gbOnlyAkk = False
            gbRunDVS = True
            giStage2 = 0
            gbAkkum = False
        End If
    End If
End Function

Public Sub InitAGNKS()   
    frmStart.tmrMotor.Interval = 65535
    frmStart.tmrMotor.Enabled = False
    

    gbCmdStart = True    'Сначала Пуск АГНКС
    giMainРасход = 1    'Начинаем добавлять к показанию ИР1
    
    InitDisk
    ConnectKKM
    Init_Controllers
    ResetExpenseCounter_1
    ResetExpenseCounter_2
End Sub

Private Sub Init_Controllers()
    'Инициализация платы ACL8113
    Init_ISO813_Driver
    'Инициализация платы Pet48DIO
    Init_DIO_Driver
End Sub

Public Function Verify_Damage()
    Dim s           As String
    'Функция проверки аварийных датчиков
    s = ""
    If gnДатчик(45).Data = 1 Then
        s = s & "Пожар в отсеке ДВС ! "
        If gbStopAGNKS = False Then
            ROff A1, 1 'закрыть К 1-6
            ROff A0, 0 'Закрыть К7, ВЫКЛ Реле 1
            ROn A1, 2  'Стоп ДВС
            giStage2 = 0
            giStage = 3    'Переход на этап Danger
            giStage1 = 0
            gbAkkum = False
            frmStart.SSCmdStart.Enabled = False
            gbCmdStart = True
            StopOutput (2)
            gbStopAGNKS = True
        End If
    End If

    If isFireTech Then
        s = s & "Пожар в техн. отсеке ! "
        If gbStopAGNKS = False Then
            s = ОстановАГНКС
        End If

    End If

    If gnДатчик(42).Data = 1 Then
        s = s & "Загазованность 20%(отсек ДВС) ! "
        If gbStopAGNKS = False Then
            s = ОстановАГНКС
        End If
    End If
    If gnДатчик(44).Data = 1 Then
        s = s & "Загазованность 20%(техн.отсек) ! "
        If gbStopAGNKS = False Then
            s = ОстановАГНКС
        End If
    End If
    If gnДатчик(40).Data = 1 Then
        s = s & "Потеря напряжения 220 В ! "
    End If

    'If gnДатчик(9).Data = 1 Then
    '  s = s & "Отказ прибора Метан в тех.отсеке ! "
    'End If
    If gnДатчик(33).Data = 1 Then
        s = s & "Высокая tC охл.жидкости ДВС ! "
    End If
    If gnДатчик(35).Data = 1 Then
        s = s & "Падение Давл. в системе смазки ДВС ! "
    End If

    If gnДатчик(41).Data = 1 Then
        s = s & "Загазованность 10%(отсек ДВС) ! "
    End If
    If gnДатчик(43).Data = 1 Then
        s = s & "Загазованность 10%(техн.отсек) ! "
    End If
    If gnДатчик(47).Data = 1 Then
        s = s & "Отказ СТМ-10 ! "
    End If

    If gnDif(13) > 60 Then
        s = s & "Повышена температура на выходе компрессора ! "
    End If

    Verify_Damage = s
End Function




