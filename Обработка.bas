Attribute VB_Name = "Module3"
Option Explicit


Public Function ОпросПлат() As String
    Dim i As Long
    Dim j As Integer
    Dim s As String
    Dim f As Double
    ОпросПлат = ""
    
    'Прочитать состояние портов платы PET-48DIO
        For i = 0 To 5
         If i = 3 Then
            'glРезультат = W_48DIO_Read_Back((i - 3) + 256, glЗначение)
            glЗначение = DIO_InputByte(glАдрес + 4)
            gn48DIO(i) = CInt(glЗначение)
         ElseIf i = 0 Then
            'glРезультат = W_48DIO_Read_Back(i, glЗначение)
            glЗначение = DIO_InputByte(glАдрес)
            gn48DIO(i) = CInt(glЗначение)
         ElseIf i < 3 Then
                'glРезультат = W_48DIO_DI(i, glЗначение)
            glЗначение = DIO_InputByte(glАдрес + i)
            gn48DIO(i) = Not (CInt(glЗначение))
         Else
                'glРезультат = W_48DIO_DI((i - 3) + 256, glЗначение)
            glЗначение = DIO_InputByte(glАдрес + 1 + i)
            gn48DIO(i) = Not (CInt(glЗначение))
         End If
        Next i
        
    'Опрос каналов платы ACL8113
    For i = 0 To 16
        For j = 0 To 3
            'glРезультат = W_8113_AD_Aquire(i, glЗначение)
            f = ISO813_AD_Float(glaАдрес, i, 1, 0, 1)
            'glaАдрес - I/O port base address
            'i - номер канала
            '1 - A/D Gain : 0 0~10 V
            '               1 0~5 V
            '0 - Unipolar
            '1 = 10V     0 = 20 V
        Next j
        
       ggACL8113(i) = f
    Next i
    
    
    If ОпросПлат = "" Then ОпросПлат = "OK"

End Function


Public Function Обработка_1()
    Dim p As Integer
    Dim r As Integer
    Dim i As Integer
    Dim j As Integer
    Dim Temp As Double
    Dim IR As Double
    Dim s As String
    'Отобразить состояние датчиков, работающих с платой Pet48DIO
    s = ""
        For i = 0 To 5
            p = gn48DIO(i)
            s = s & CStr(p) & " "
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
    
    frmStart.lblPC.Caption = s
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
          If (gnDif(i) - 4) <= 17 And (gnDif(i) - 4) >= -1 Then
          Select Case (i)
          Case 8
            gnDif(i) = 200 * ((Temp + 1) / 18) - 50
          Case 9
            gnDif(i) = 12.5 * (gnDif(i) - 4) - 50
          Case 12
            gnDif(i) = 6.25 * (gnDif(i) - 4) - 50
            Case 13
            gnDif(i) = 150 * ((Temp + 1) / 18) - 50
            Case Else
            gnDif(i) = 6.25 * (gnDif(i) - 4) - 50
            End Select
          Else
            gnDif(i) = -1
          End If
        Next i

    
    'Пересчёт измеренных значений в давление (в МПа):
    
        For i = 4 To 7
        'Проверка (i - 4) , если не удовлетворяет то -1
          If (gnDif(i) - 4) <= 17 And (gnDif(i) - 4) >= -1 Then
            gnDif(i) = gKp_1 * (gnDif(i) - 4)
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
        If (gnDif(i) - 4) <= 17 And (gnDif(i) - 4) >= -1 Then
          gnDif(i) = (gnDif(i) - 4) * gKp
        Else
          gnDif(i) = -1
        End If
      Next i
      
     ' Расчет оборотов ДВС
        If (gnDif(14) - 4) <= 17 And (gnDif(14) - 4) >= -1 Then
          gnDif(14) = (gnDif(14) - 4) * 200 ' =3200/16
        Else
          gnDif(14) = -1
        End If
   Call AddSensorsData(2, gnDif(5), gnDif(11), gnDif(4), 1.5, 0.95 * gdK, 0)
      gdИР2 = GetMass(2)
      Temp = GetMassExpense(2)
   'Считать расход (общий) по ИР1
   If giMainРасход = 1 Then
     Call AddSensorsData(1, gnDif(2), gnDif(9), gnDif(3), 6, 0.95 * gdK, 0)
      IR = GetMass(1)
      gdИР1 = gdInitIR + IR
   Else
     Call AddSensorsData(1, gnDif(2), gnDif(9), gnDif(3), 6, 0.95 * gdK, -(Temp))
     IR = GetMass(1)
     gdИР1 = gdInitIR + IR
   End If

End Function


