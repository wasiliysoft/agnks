Attribute VB_Name = "Module4"

Public Function Convert_Date(ss As String)
Dim s2 As String
Dim i As Integer
s2 = "#"
   For i = 0 To Len(ss)
     If Mid(ss, i + 1, 1) = "." Then
        s2 = s2 & "/"
     Else
        s2 = s2 + Mid(ss, i + 1, 1)
     End If
    Next i
 Convert_Date = s2 & "#"
End Function

'Функция обработки аварийных ситуаций
Public Function Danger() As String
If gnДатчик(29).Data = 1 Then
  'Если перепад во вх. и вых. рукове станет меньше 5 кг
  If Abs(gnDif(6) - gnDif(2)) <= 0.5 Then
    ROff A1, 223 'Закрыть КЭМ4
    If gbFireTech = True Then
      ROn A1, 24 'И открыть КЭМ3 и КЭМ2
      ROn A0, 2  'открыть КЭМ7
    Else
      ROn A1, 16 'И открыть КЭМ3
    End If
  End If
End If
frmStart.cmdDanger.Visible = True
End Function

'Функция инициализации данных, считывание с диска
Public Function InitDisk() As Integer
'Возвращает код ошибки : 0 -нет ошибок
Dim i As Integer
Dim j As Integer
Dim j1 As Integer

Dim k As Integer
Dim t As MyRecType
Dim temp1 As MyRecType
Dim temp2 As MyRecType
Dim d As Date
Dim sum As Double
Dim idx1, idx2, idx3 As Integer
'Для ввода поправочного коэффициента
Dim descr As Integer
Dim sPath As String
Dim rec As pswd
Dim s As String
Dim s1 As String
Dim v As Variant
On Error Resume Next
'Открыть все файлы

frmStart.MousePointer = vbHourglass
s = App.Path + "\base.mdb"
Set StatWS = DBEngine.Workspaces(0)
Set StatDB = StatWS.OpenDatabase(s)
Set StatRS = StatDB.OpenRecordset("stat", dbOpenTable)

If Not StatRS.EOF Then
'Если не пустая база данных
' StatRS.MoveFirst
 'Получаем самую раннюю дату
 Set SelectRS = StatDB.OpenRecordset("select MIN(DATA) from stat ")
 temp1.dt = SelectRS(0)
 'Получаем самую позднюю дату
' StatRS.MoveLast
 Set SelectRS = StatDB.OpenRecordset("select MAX(DATA) from stat ")
 temp2.dt = SelectRS(0)
 s = Convert_Date(Str(Month(temp2.dt)) & "/" & Day(temp2.dt) & "/" & Year(temp2.dt) & " " & Hour(temp2.dt) & ":" & Minute(temp2.dt) & ":" & Second(temp2.dt))
 
 Set SelectRS = StatDB.OpenRecordset("SELECT * From stat WHERE stat.data=" & s)

  GMC = SelectRS("MOTO")
 gDateRec = Now  ' temp2.dt

 'temp2.Motor = SelectRS("MOTO")
 d = Now
 
 'Для заполнения 4 колонки в цикле посчитать все суммарные заправки за годы начиная с самого раннего и до предыдущего
 For i = Year(temp1.dt) To Year(d) - 1
  s = "1/1/" + CStr(i)
  s = Convert_Date(s)
  s1 = "12/31/" + CStr(i)
  s1 = Convert_Date(s1)
    Set SelectRS = StatDB.OpenRecordset("select SUM(GAZ_CAR) from stat where (stat.DATA between " & s & " AND " & s1 & ")")
  If IsNull(SelectRS(0)) Then
  Else
    s = Format(CStr(i)) + "        " + Format(CStr(SelectRS(0)), "###0.00")
    frmStart.lstStat(3).AddItem (s)

  End If
 Next i
 'Для заполнения 3 колонки в цикле посчитать все суммарные заправки за месяцы начиная с января этого года и до пред.
If Year(temp2.dt) = Year(d) Then
 For i = 1 To Month(d) - 1
  s = CStr(i) + "/1/" + CStr(Year(d))
  s = Convert_Date(s)
  
   Select Case i
   Case 1, 3, 5, 7, 8, 10, 12
       s1 = CStr(i) + "/31/" + CStr(Year(d))
   Case 2
       Select Case Year(d)
       Case 2004, 2008, 2012, 2016, 2020, 2024, 2028, 2032, 2036, 2040
         s1 = CStr(i) + "/29/" + CStr(Year(d))
       Case Else
          s1 = CStr(i) + "/28/" + CStr(Year(d))
       End Select
   Case Else
       s1 = CStr(i) + "/30/" + CStr(Year(d))
   End Select
   
  s1 = Convert_Date(s1)
    Set SelectRS = StatDB.OpenRecordset("select SUM(GAZ_CAR) from stat where (stat.DATA between " & s & " AND " & s1 & ")")
  v = SelectRS(0)
  If IsNull(v) Then
  Else
    s = Format(CStr(i), "00") + "        " + Format(CStr(SelectRS(0)), "###0.00")
    frmStart.lstStat(2).AddItem (s)
 End If
 Next i
End If
 'Для заполнения 2 колонки в цикле посчитать все суммарные заправки за дни начиная с 1 этого месяца и до пред.
 
 frmStart.lblStat(0).Caption = "За " + Format(d, "dd")
 frmStart.lblStat(1).Caption = "За " + Format(d, "mmmm")
 frmStart.lblStat(2).Caption = "За " + Format(d, "yyyy")

 If Month(temp2.dt) = Month(d) Then
 For i = 1 To Day(d) - 1
  s = CStr(Month(d)) + "/" + CStr(i) + "/" + CStr(Year(d))
  s = Convert_Date(s)
  s1 = CStr(Month(d)) + "/" + CStr(i + 1) + "/" + CStr(Year(d))
  s1 = Convert_Date(s1)

    Set SelectRS = StatDB.OpenRecordset("select SUM(GAZ_CAR) from stat where (stat.DATA between " & s & " AND " & s1 & ")")
  v = SelectRS(0)
  If IsNull(v) Then
  Else
   s = Format(CStr(i), "00") + "        " + Format(CStr(SelectRS(0)), "###0.00")
   frmStart.lstStat(1).AddItem (s)
  End If
 Next i
 End If
 
 'Если дата последней заправки = текущей, то в 1 колонку занести заправки за эту дату.
  s = Format(d, "mm/dd/yyyy")
  s = Convert_Date(s)
  s1 = Format(d + 1, "mm/dd/yyyy")
  s1 = Convert_Date(s1)
  'frmShow.MousePointer = vbHourglass
Set SelectRS = StatDB.OpenRecordset("select * from stat where DATA between " & s & " AND " & s1)
If SelectRS.RecordCount >= 1 Then
    SelectRS.MoveLast
    SelectRS.MoveFirst
    
 For i = 0 To SelectRS.RecordCount - 1
   s = ""
   s = Format(CStr(SelectRS("Data")), "hh:mm:ss") + "        " + Format(CStr(SelectRS("GAZ_CAR")), "###0.00")
   frmStart.lstStat(0).AddItem (s)
   SelectRS.MoveNext
 Next i
End If

 'GMC присвоить SelectRS("MOTO") последней по дате записи
 'gDateRec присвоить SelectRS("Data")последней по дате записи
 '
  'gdInitIR = 0
  'd = Now
  'sum = 0
  'k = Int(gdaStat1(0).IR1) 'количество записей в массиве 1
  'gDateRec = gdaStat1(1).dt
  'GMC = gdaStat1(k).Motor

'RefreshStat

 'Call showfields
Else
 GMC = 0
 gDateRec = Now  ' temp2.dt

End If

gDateRec = Now
sPath = "C:\Winnt\dll32.dll"
descr = FreeFile
Open sPath For Random As descr Len = Len(rec)
If FileLen(sPath) = 0 Then
  rec.pwd = "LAB"
  rec.PC = 1
  Put #descr, 1, rec
  Password = "LAB"
  gdK = rec.PC
Else
  Get #descr, 1, rec
  Password = Trim(rec.pwd)
  gdK = rec.PC
End If
  Close #descr
  
InitDisk = 0
  

frmStart.MousePointer = vbArrow
gdPlot = 0.7
Dim fh As Long
fh = FreeFile
s = App.Path & "\price.txt"
Open s For Input Access Read As fh
Seek #fh, 1
Line Input #fh, s
gdPrice = CDbl(s)
Line Input #fh, s
gdPlot = CDbl(s)
Close #fh
If gdPlot < 0.5 Then gdPlot = 0.7
If gdPlot > 1 Then gdPlot = 0.7
frmStart.Caption = frmStart.Caption & " Плотность газа = " & CStr(gdPlot) & " кг/м3"
Exit Function

ErrorHandler:      'Если есть какие-нибудь ошибки возвращаем -1
  InitDisk = -1
Exit Function
  
End Function

'Инициируется табло
Public Sub InitTablo()
giaTableDecoder(1) = 243 'Соответствует светящейся 1
giaTableDecoder(2) = 73  'Соответствует светящейся 2
giaTableDecoder(3) = 97  'Соответствует светящейся 3
giaTableDecoder(4) = 51  'Соответствует светящейся 4
giaTableDecoder(5) = 37  'Соответствует светящейся 5
giaTableDecoder(6) = 5   'Соответствует светящейся 6
giaTableDecoder(7) = 241 'Соответствует светящейся 7
giaTableDecoder(8) = 1   'Соответствует светящейся 8
giaTableDecoder(9) = 33  'Соответствует светящейся 9
giaTableDecoder(0) = 129 'Соответствует светящейся 0
giaTableDecoder(10) = 255 ' Cоответствует погашенному индикатору
End Sub

'Процедура обновляет информацию в окне статистики
Public Sub RefreshStat()
Dim i As Integer
Dim s As String
  For i = 0 To 3
     frmStart.lstStat(i).Clear
  Next i
  
  If gdaStat1(1).dt = 0 Then
    frmStart.lblStat(0).Caption = "За " & Format(Now, "d")
  Else
    frmStart.lblStat(0).Caption = "За " & Format(gdaStat1(1).dt, "d")
  End If
  
  If gdaStat2(1).dt = 0 Then
    frmStart.lblStat(1).Caption = "За " & Format(Now, "mmmm")
  Else
    frmStart.lblStat(1).Caption = "За " & Format(gdaStat2(1).dt, "mmmm")
  End If

  If gdaStat3(1).dt = 0 Then
    frmStart.lblStat(2).Caption = "За " & Format(Now, "yyyy")
  Else
    frmStart.lblStat(2).Caption = "За " & Format(gdaStat3(1).dt, "yyyy")
  End If

  
  For i = 1 To gdaStat1(0).IR1
     s = Format(gdaStat1(i).dt, "hh:mm:ss") + "     " + Format(gdaStat1(i).IR2, "###0.00")
     frmStart.lstStat(0).AddItem s
  Next
  
  For i = 1 To gdaStat2(0).IR1
     s = Format(gdaStat2(i).dt, "mmmm d yyyy") + "     " + Format(gdaStat2(i).IR2, "###0.00")
     frmStart.lstStat(1).AddItem s
  Next
  
  For i = 1 To gdaStat3(0).IR1
     s = Format(gdaStat3(i).dt, "mmmm ") + "     " + Format(gdaStat3(i).IR2, "###0.00")
     frmStart.lstStat(2).AddItem s
  Next
  
  For i = 1 To gdaStat4(0).IR1
     s = Format(gdaStat4(i).dt, "yyyy") + "     " + Format(gdaStat4(i).IR2, "###0.00")
     frmStart.lstStat(3).AddItem s
  Next

End Sub

'Функция выводит в port 0
Public Function ROff(port As Integer, n As Integer) As Integer
Dim t As Byte
Dim j As Integer

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
  Case Else   '''Возможно и не надо !!!
     j = 3
End Select

If j <= 2 Then
     t = gn48DIO(j) 'считываем состояние порта
Else
    t = gn48DIO(j - 1) 'считываем состояние порта
End If
        'Закрыть
     t = t And n ' 0 в n-ый канал
        ''''Для отладки !!!
     'W_48DIO_DO port, t
     DIO_OutputByte glАдрес + j, t
     
     gn48DIO(j) = t
     ОпросПлат 'Нужно чтобы узнать ИР2
      '----------
          '----------
End Function

'Функция выводит в port 1
Public Function ROn(port As Integer, n As Integer) As Integer
Dim t As Byte
Dim j As Integer

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
  Case Else   '''Возможно и не надо !!!
     j = 3
End Select
If j <= 2 Then
     t = gn48DIO(j) 'считываем состояние порта
Else
    t = gn48DIO(j - 1) 'считываем состояние порта
End If
        'Открыть
     t = t Or n ' 1 в n-ый канал
        ''''Для отладки !!!
'     W_48DIO_DO port, t
      't = 2
    DIO_OutputByte glАдрес + j, t

     gn48DIO(j) = t
          '----------
     ОпросПлат
End Function

'Функция остановки АГНКС
Public Function ОстановАГНКС() As String
Dim s As String
   'закрыть все КЭМы
   ROff A1, 1
   'Стоп ДВС, открыть КЭМ2, открыть КЭМ4
   'если загазованность 20 %
   If (gnДатчик(42).Data = 1) Or (gnДатчик(44).Data = 1) Then
   'Стоп ДВС, открыть КЭМ2, открыть КЭМ4
     ROn A1, 42
     ROn A0, 2 'Открыть КЭ7
   ElseIf gnДатчик(46).Data = 1 Then 'если пожар в техническом отсеке
     ROn A1, 34 'стоп ДВС, открыть КЭМ4
   End If
     
   
     giStage2 = 0
     giStage = 3 'Переход на этап Danger
     giStage1 = 0
     gbAkkum = False
     frmStart.SSCmdStart.Enabled = False
     gbCmdStart = True
     frmStart.SSCmdStart.Caption = "Пуск АГНКС"
     gbDVSStopping = True
   ОстановАГНКС = "Останов АГНКС"
End Function

'Функция остановки ДВС
Public Function ОстановДВС() As String
 'Если открыт КЭМ5 - закрыть
    If gnДатчик(30).Data = 1 Then
      ROff A1, 191
    End If
      'Открыть Кэм4
     ROn A1, 32
  
     giStage2 = 0
     giStage = 1 'Переход на этап ИсхСост
     giStage1 = 1
     giMainРасход = 0

     gbAkkum = False
     frmStart.SSCmdStart.Enabled = False
     gbCmdStart = False
     gbDVSStopping = True
End Function
'Сама процедура заправки
Public Function Заправка()
Dim dFullCar As Double 'Здесь запоминаем давление в баке автомобиля
Dim s, s1 As String
Dim MaxIR As Double 'Запоминаем max расход при открытии КЭМ6
Dim p As Double

   ' ПОДЭТАП 8  - Заправка только от ов
 If giStage2 = 8 Then
    'Заправка машин
    
        If gnДатчик(18).Data = 1 Then
       'Закрыть КЭ4
          ROff A1, 223
        Else
       'Закрыть КЭ3
          ROff A1, 239
        End If
     
      
     ROn A1, 64         'Открыть КЭ5
     giStage2 = 9
     ROn A1, 128
     gdРасход1 = 0 'Обнуляем расход на одну машину
     ResetExpenseCounter (2)
     StartOutput (2)
     gbDontStat = True 'Нельзя работать с диском
     Exit Function
 End If
 
 'ПодЭтап 9
 If giStage2 = 9 Then
      If (Abs(gnDif(5) - gnDif(4)) > 0.5) Then
        Заправка = "Идет заправка "
        'Считаем расход на одну машину (за полсекунды)
        gdTime = GetTimeCounter(2)
        gdРасход1 = gdИР2
        Exit Function
      Else
        'Закрыть пистолет
            'Закрыть КЭ5
            ROff A1, 191
             ROff A1, 127
             StopOutput (2)
            gbDontStat = False 'Можно работать с диском
            
        gdTime = GetTimeCounter(2)
        
            'Заполнить статистику по заправке
           
           '<<<<Прекратить считать расход>>>>
           StatRS.AddNew
  
           StatRS("DATA") = Now
           StatRS("GAZ_CAR") = gdРасход1 / gdPlot '* 1.42
 
           StatRS("GAZ_IR1") = gdИР1
           StatRS("MOTO") = GMC + MotorCount
           GMC = GMC + MotorCount
            MotorCount = 0
   If (Day(gDateRec) < Day(Now)) Or (Month(gDateRec) < Month(Now)) Or (Year(gDateRec) < Year(Now)) Then
                 Verify
              End If
          
           StatRS.Update

     s = Format(Now, "hh:mm:ss") + "        " + Format((gdРасход1 / gdPlot), "###0.00")
     frmStart.lstStat(0).AddItem s
           
            gDateRec = Now
            
            gbЗаправка = False
           
        'Разрешить повторную заправку автомобиля во время заправки аккумуляторов
          frmStart.SSCmdStart.Enabled = True
          gbAkkum = True
          giStage = 1  'Переход на Этап Предпуска
          giStage1 = 0
          giStage2 = 0
       Exit Function
    
       End If
   End If
 

 ' ПОДЭТАП 1
  If (giStage2 = 0) And (gbFrmShow = False) Then
   gsMsg = "Пистолет вставлен ?"
   frmЗапрос.Show 0
     gbFrmShow = True
   Заправка = "Выведен диалог (пистолет)"
   Exit Function
  End If
  
 ' ПОДЭТАП 2
  If (giStage2 = 1) And (gbFrmShow = False) Then
    If giTrigger = 0 Then
        giStage2 = 0
        gbЗаправка = True
        gbAkkum = False
      giStage = 1 'Переход на этап ПредПуск
      giStage1 = 1 'Сразу на проверку ДВС и компрессора
      
      Заправка = "Переход на этап ПредПуск"
      Exit Function
    Else
      giStage2 = 2
      Car = 1
      s = "Заправка машин"
    End If
  End If
  
   ' ПОДЭТАП 3
  If giStage2 = 2 Then
    If Car = 1 Then
      'Заправка машин
      
     ROn A1, 64         'Открыть КЭ5
     giStage2 = 3
     gdРасход1 = 0 'Обнуляем расход на одну машину
     ResetExpenseCounter (2)
     StartOutput (2)
     gbDontStat = True 'Нельзя работать с диском
     giMainРасход = 1
    End If
  End If
    
  ' ПОДЭТАП 4
  If giStage2 = 3 Then
       dFullCar = gnDif(5) 'Запоминаем давление в баке машины
  'Считать расход заправки автомобиля
       gbЗаправка = True

       If (gnDif(7) - dFullCar) >= 2 Then 'Разница давлений в аккумуляторах и баке
          'Открыть КЭ6 - заправка и от аккумуляторов
          ROn A1, 128
       End If
       
        If gnДатчик(18).Data = 1 Then
       'Закрыть КЭ4 - Начинает гнать газ компрессор
          ROff A1, 223
        Else
       'Закрыть КЭ3 - Начинает гнать газ компрессор
          ROff A1, 239
        End If
        
       giStage2 = 4 'Переходим к подэтапу заправки аккумуляторов
       ОпросПлат
       Обработка_1
  End If
    
  ' ПОДЭТАП 5
   If giStage2 = 4 Then
     '<<<<Считать расход по ИР2>>>>
                                                    
                                                    
        MaxIR = GetMassExpense(2)
        If (gbAkkum = False) And (((gnДатчик(20).Data = 1) And (((MaxIR * 3600) <= gdRashAkkEnd) _
        And (MaxIR > 0)) And (GetTimeCounter(2) >= 5)) Or ((gnDif(7) - gnDif(4)) <= 0.5)) Then
            'Закрыть КЭ6
           ROff A1, 127
           'Exit Function
        End If
 ' Убрал проверку на остановку заправки по минимальному расходу - отключает раньше
'If (gbAkkum = False) And ((Not (gnDif(4) >= gdUpLevel)) And (MaxIR * 3600 >= 35)) Then
If (gbAkkum = False) And ((Not (gnDif(4) >= gdUpLevel))) Then

        Заправка = "Идет заправка "
        'Считаем расход на одну машину (за полсекунды)
        gdРасход1 = gdИР2
        
        gdTime = GetTimeCounter(2)
        Exit Function
      ElseIf (gbAkkum = False) Then
        'Закрыть пистолет
            'Закрыть КЭ5
            ROff A1, 191
            gbDontStat = False 'Можно работать с диском
            StopOutput (2)
        gdTime = GetTimeCounter(2)
        
            'Заполнить статистику по заправке
                       StatRS.AddNew
           StatRS("DATA") = Now
           StatRS("GAZ_CAR") = gdРасход1 / gdPlot '* 1.42
 
           StatRS("GAZ_IR1") = gdИР1
           StatRS("MOTO") = GMC + MotorCount
           GMC = GMC + MotorCount
            MotorCount = 0
   If (Day(gDateRec) < Day(Now)) Or (Month(gDateRec) < Month(Now)) Or (Year(gDateRec) < Year(Now)) Then
     Verify
   End If
            
           StatRS.Update

     s = Format(Now, "hh:mm:ss") + "        " + Format((gdРасход1 / gdPlot), "###0.00")
     frmStart.lstStat(0).AddItem s

           
           gDateRec = Now

           '<<<<Прекратить считать расход>>>>
            gbЗаправка = False
           
           'ЗАПРАВЛЯЕМ АККУМУЛЯТОРЫ
           'Открыть КЭ6
           ROn A1, 128
             
      'Разрешить повторную заправку автомобиля во время заправки аккумуляторов
           frmStart.SSCmdStart.Enabled = True
           gbAkkum = True
       End If
       
    'заправлять аккумуляторы до 200 кг
    If (gnDif(7) < gdUpLevel) And (gbAkkum = True) Then
      Заправка = "Заправка аккумуляторов"
      Exit Function
    Else
        'Закрыть КЭ6
        ROff A1, 127
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
        'Открыть КЭ5
     ROn A1, 64
     dFullCar = gnDif(5) 'Запоминаем давление в баке машины
     s = "Переходим на заправку машин"
     ResetExpenseCounter (2)
     StartOutput (2)
       ОпросПлат
       Обработка_1

     giStage2 = 4
  'Считать расход заправки автомобиля
     giMainРасход = 1
  gbЗаправка = True
  gbAkkum = False
  gdРасход1 = 0 'Обнуляем расход на одну машину
  
  gdTime = GetTimeCounter(2)
 End If
 
 
Заправка = s
End Function


'Приводит АГНКС в исходное состояние
Public Function ИсхСост() As String
Dim s As String
Dim norma As Boolean
 frmStart.SSCommand2(1).Enabled = True
 frmStart.SSCommand2(0).Enabled = True
 gbFireDVS = False
 gbFireTech = False
     s = ""
     norma = True
     gbRunDVS = False
   'Входные реле включены (порт A0 и A1) ? - неисправны
   If (gnДатчик(16).Data = 1) Or (gnДатчик(17).Data = 1) Or (gnДатчик(18).Data = 1) Or _
   (gnДатчик(19).Data = 1) Or (gnДатчик(20).Data = 1) Or (gnДатчик(21).Data = 1) Then
     s = "Есть открытые КЭМы !!!"
     norma = False
   End If
   
   If (gnДатчик(25).Data = 1) Then
     s = s & "Нажата Останов ДВС !!!"
     norma = False
   End If
   
  ' If gnDif(1) <= 0.3 Then   '0.3 потому что плавает показание
  '   s = s & "Нет газа !!!"
  '   norma = False
  ' End If
   
   If gnДатчик(36).Data = 1 Then
     s = s & "Включена муфта !!!"
     norma = False
   End If
   
   If gnDif(14) > 100 Then   'Есть обороты
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
      'Открыть КЭ4
        ROn A1, 32
      Else
     'Открыть КЭ3 - для запуска ДВС
        ROn A1, 16
     End If
     gbAkkum = True
     giStage1 = 1 ' Пререход на второй подэтап
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
        giStage = 0 'Переход на этап ИсхСост
        giStage1 = 0
        gbAkkum = False
        gbRunDVS = False
        frmStart.SSCmdStart.Enabled = False
        gbCmdStart = True
        frmStart.SSCmdStart.Caption = "ПУСК АГНКС"
        gbDVSStopping = True
        'frmStart.Timer2.Enabled = False
    'Закрыть все Кэм
        If (gnДатчик(16).Data = 1) Or (gnДатчик(17).Data = 1) Or (gnДатчик(18).Data = 1) Or _
           (gnДатчик(19).Data = 1) Or (gnДатчик(20).Data = 1) Or (gnДатчик(21).Data = 1) Or _
           (gnДатчик(25).Data = 1) Then
           ROff A1, 1
        End If

     End If
     
     If gnДатчик(36).Data = 0 Then
       ПредПуск = "Двигатель готов к запуску !!!"
       Exit Function
     Else
       ПредПуск = "Двигатель не готов к запуску !!! Включена муфта "
       Exit Function
     End If
   'Есть обороты ДВС
   ElseIf gnДатчик(36).Data = 0 Then
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
Dim err As Integer
   gKv = 1     'масштабный коэффициент пересчёта в напряжение
                        'для однополярного входа на 10в
                    
    gKi = gKv / 0.2   'масштабный коэффициент пересчёта в ток (миллиамперы)
                        'при сопротивлении нагрузки 0,448 кОм
                        
    gKi_1 = gKv / 0.2   'масштабный коэффициент пересчёта в ток (миллиамперы)
                          'при сопротивлении нагрузки 0,187 кОм

    gKp = 1.6 / 16    'масштабный коэффициент пересчёта тока в давление

    gKp_1 = 25 / 16    'масштабный коэффициент пересчёта тока в давление
    
    gKn = (2740 + 448) / 448 'Коэффициент пересчета рабочего напряжения на датчиках
    
    gdUpLevel = 200 * 0.0981 'Предел давления для заправки
    'frmStart.label5.Caption = "196"
    giChanel = 16 'Номер канала дл напряжения
    MotorCount = 0
    frmStart.tmrMotor.Interval = 65535
    frmStart.tmrMotor.Enabled = False
    
  ' Инициализация плат
   giDVS = 0
   giMAX = 0

    giStage = 0  'Инициализация процедуры заправки (Переходим на ИсхСост)
    giStage1 = 0
    giStage2 = 0
    gbFrmShow = False
    frmStart.SSTab1.Tab = 3
    gbCmdStart = True 'Сначала Пуск АГНКС
    gbAkkum = False
    Car = 0
     glAver = 0
     glCounter = 0
     gbDVSStopping = False
    gbRunDVS = False
     DVSEmul = False
     MFTEmul = False
     gbHahdControl = False
    gdРасход1 = 0
    
    gdRashAkkEnd = 65
    gdK = 1
    '+++++++++++++++++
    gdAllРасход = 0 ' Пока обнуляю, а в будующем нужно запоминать в файле и считывать
    '+++++++++++++++++
    
    gbDontStat = False
      
   gbStopAGNKS = False
   giMainРасход = 1 'Начинаем добавлять к показанию ИР1
   giOnlyAkk = False
 err = InitDisk()
  If err = -1 Then
     'Обработка ошибки работы с диском
     'Записать все данные на диск
     'Если была ошибка связанная с журналом, то обнуляем его
End If
' Для отладки !!!!!!!(отключить)
ConnectKKM
InitTablo
s = ИниКонтроль()
ResetExpenseCounter (1)
ResetExpenseCounter (2)
End Sub

Public Sub ShowPict()
Dim s As String
With frmStart
'Управление кранами
  If gnДатчик(16).Data = 1 Then
    .КЭ2(0).Visible = False
    .КЭ2(1).Visible = True
    .Факел(0).Visible = True
  Else
    .КЭ2(0).Visible = True
    .КЭ2(1).Visible = False
    .Факел(0).Visible = False
  End If

  If gnДатчик(17).Data = 1 Then
    .КЭ3(0).Visible = False
    .КЭ3(1).Visible = True
  Else
    .КЭ3(0).Visible = True
    .КЭ3(1).Visible = False
  End If

  If gnДатчик(18).Data = 1 Then
    .КЭ4(0).Visible = False
    .КЭ4(1).Visible = True
  Else
    .КЭ4(0).Visible = True
    .КЭ4(1).Visible = False
  End If
  
  If gnДатчик(19).Data = 1 Then
    .КЭ5(0).Visible = False
    .КЭ5(1).Visible = True
  Else
    .КЭ5(0).Visible = True
    .КЭ5(1).Visible = False
  End If
  
  If gnДатчик(20).Data = 1 Then
    .КЭ6(0).Visible = False
    .КЭ6(1).Visible = True
  Else
    .КЭ6(0).Visible = True
    .КЭ6(1).Visible = False
  End If
  
  If gnДатчик(23).Data = 1 Then
    .КЭ7(0).Visible = False
    .КЭ7(1).Visible = True
  Else
    .КЭ7(0).Visible = True
    .КЭ7(1).Visible = False
  End If
  
  If gnДатчик(21).Data = 1 Then
    .КЭ1(0).Visible = False
    .КЭ1(1).Visible = True
  Else
    .КЭ1(0).Visible = True
    .КЭ1(1).Visible = False
  End If
  
  'Управление муфтой
  If gnДатчик(36).Data = 1 Then
    .Муфта.BackColor = &HFF&
  Else
    .Муфта.BackColor = &HC0C0C0
  End If
  
  'Повышение температуры охлаждения ДВС
  If gnДатчик(33).Data = 1 Then
    gnТ_ДВС = 1
  Else
    gnТ_ДВС = 0
  End If
  
  'Пожар в отсеке ДВС
  If gnДатчик(45).Data = 1 Then
    gnПожар_ДВС = 1
  Else
    gnПожар_ДВС = 0
  End If

  'Пожар в технологическом отсеке
  If gnДатчик(46).Data = 1 Then
    gnПожар_Тех = 1
  Else
    gnПожар_Тех = 0
  End If
 
  'Газ в отсеке ДВС
  If (gnДатчик(41).Data = 1) Then
    gnГаз_10_ДВС = 1
  Else
    gnГаз_10_ДВС = 0
  End If
  
  If (gnДатчик(42).Data = 1) Then
    gnГаз_20_ДВС = 1
  Else
    gnГаз_20_ДВС = 0
  End If
  
  'Газ в технологическом отсеке
  If (gnДатчик(43).Data = 1) Then
    gnГаз_10_Тех = 1
  Else
    gnГаз_10_Тех = 0
  End If
  
  If (gnДатчик(44).Data = 1) Then
    gnГаз_20_Тех = 1
  Else
    gnГаз_20_Тех = 0
  End If
 End With
 
  s = Format(gdK, "0.000")
  s = s + "   - коэффициент"
  frmStart.lblPC.Caption = s


End Sub


'Процедура проверки переходов дат
Public Function Verify()
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim d As Date
Dim sum1 As Double
Dim sum2 As Double
Dim sum3 As Double
Dim Old As String
Dim s As String
Dim s1 As String
 
'Проверка перехода даты
d = Now
sum1 = 0
sum2 = 0
sum3 = 0
'If gDateRec < d Then
 '           MsgBox ("Verify" + Str(gDateRec) + Str(d))

 If gDateRec < d Then
     
   For i = 0 To frmStart.lstStat(0).ListCount - 1
     s1 = frmStart.lstStat(0).List(i)
     s = Mid(s1, 17, Len(frmStart.lstStat(0).List(i)) - 1)
     sum1 = sum1 + CDbl(s)
   Next i
   'Строка послдней заправки
    Old = s1
   frmStart.lblStat(0).Caption = "За " + Format(d, "dd")
 frmStart.lblStat(1).Caption = "За " + Format(d, "mmmm")
 frmStart.lblStat(2).Caption = "За " + Format(d, "yyyy")
  
   frmStart.lstStat(0).Clear
  If (Month(gDateRec) < Month(d)) Or ((Month(gDateRec) > Month(d)) And (Year(gDateRec) < Year(d))) Then
   For i = 0 To frmStart.lstStat(1).ListCount - 1
     s1 = frmStart.lstStat(1).List(i)
     s = Mid(s1, 11, Len(frmStart.lstStat(1).List(i)) - 1)
     sum2 = sum2 + CDbl(s)
   Next i
   frmStart.lstStat(1).Clear
   If (sum1 + sum2) <> 0 Then
    s = Format(CStr(Month(gDateRec)), "00") + "        " + Format(CStr(sum2 + sum1), "###0.00")
    frmStart.lstStat(2).AddItem (s)
   End If
  ElseIf (Month(gDateRec) = Month(d)) And (sum1 <> 0) Then
    s = Format(CStr(Day(d - 1)), "00") + "       " + Format(CStr(sum1), "###0.00")
    frmStart.lstStat(1).AddItem (s)
  End If
  
  If Year(gDateRec) < Year(d) Then
   For i = 0 To frmStart.lstStat(2).ListCount - 1
     s1 = frmStart.lstStat(2).List(i)
     s = Mid(s1, 11, Len(frmStart.lstStat(2).List(i)) - 1)
     sum3 = sum3 + CDbl(s)
   Next i
   frmStart.lstStat(2).Clear
   s = Format(CStr(Year(gDateRec)), "00") + "       " + Format(CStr(sum3), "###0.00")
   frmStart.lstStat(3).AddItem (s)
  ElseIf (Year(gDateRec) = Year(d)) And (sum2 <> 0) Then
    's = Format(CStr(Month(gDateRec)), "00") + "       " + Format(CStr(sum2), "###0.00")
   
'    s = Format(CStr(Month(d)), "00") + "       " + Format(CStr(sum2), "###0.00")
    'frmStart.lstStat(2).AddItem (s)
  End If
   gDateRec = Now
 
 End If
'End If
 
 
 
 
 'GMC = temp2.Motor
' If Not StatRS.EOF Then
'    gDateRec = temp2.dt
' Else
'    gDateRec = Now  ' temp2.dt
' End If
' frmStart.MousePointer = vbArrow

  'gdInitIR = 0
  'd = Now
  'sum = 0
  'k = Int(gdaStat1(0).IR1) 'количество записей в массиве 1
  'gDateRec = gdaStat1(1).dt
  'GMC = gdaStat1(k).Motor

'RefreshStat

 'Call showfields
End Function



Public Function Verify_Damage()
Dim s As String
'Функция проверки аварийных датчиков
s = ""
If gnДатчик(45).Data = 1 Then
  s = s & "Пожар в отсеке ДВС ! "
 If gbStopAGNKS = False Then
   
   'закрыть все КЭМы
   ROff A1, 1
   ROff A0, 0
   'Стоп ДВС
   ROn A1, 2
    gbFireDVS = True
    giStage2 = 0
     giStage = 3 'Переход на этап Danger
     giStage1 = 0
     gbAkkum = False
     frmStart.SSCmdStart.Enabled = False
     gbCmdStart = True
     frmStart.SSCmdStart.Caption = "Пуск АГНКС"
     frmStart.SSCmdStart.Visible = True
     gbDVSStopping = True
     StopOutput (2)
    gbStopAGNKS = True
 End If
End If

If gnДатчик(46).Data = 1 Then
  s = s & "Пожар в техн. отсеке ! "
  
 If gbStopAGNKS = False Then
    gbFireTech = True
    s = ОстановАГНКС
    gbStopAGNKS = True
    StopOutput (giPost)
  End If

End If

If gnДатчик(42).Data = 1 Then
  s = s & "Загазованность 20%(отсек ДВС) ! "
  If gbStopAGNKS = False Then
    s = ОстановАГНКС
    gbStopAGNKS = True
    StopOutput (2)
  End If
End If
If gnДатчик(44).Data = 1 Then
  s = s & "Загазованность 20%(техн.отсек) ! "
  If gbStopAGNKS = False Then
    s = ОстановАГНКС
    gbStopAGNKS = True
    StopOutput (2)
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


