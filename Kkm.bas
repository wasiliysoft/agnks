Attribute VB_Name = "KKM"
Option Explicit
'ConnectKKM - выполняется при загрузке основной формы



Public giErrorKKM   As Integer  'Код ошибки при выполнении функции работы с драйвером ККМ
Public gsErrorKKM   As String
Public pass         As Long
Public gsРежимККМ   As String
Public giПодРежимККМ As Integer
Public glpassKKM    As Long     'Пароль операторов ККМ при записи в таблицу ККМ
Public DrvFR        As Object   'Описание объекта драйвера ФР

'  Public Drvfr As Object
Public Sub StatusKKM()
    On Error GoTo err
    'Функция запроса статуса ККМ вызывается в начале работы программы после установки соединения
    'и далее каждые три секунды из процедуры PrintCheckKKM при выполнении операций печати и ожидания печати
    DrvFR.Password = pass     'Указываем пароль для определения статуса ККМ
    'DrvFR.GetECRStatus 'Запрос статуса ККМ
    DrvFR.GetShortECRStatus     'Запрос статуса ККМ
    giErrorKKM = DrvFR.ResultCode
    gsErrorKKM = DrvFR.ResultCodeDescription
    gsРежимККМ = DrvFR.ECRModeDescription
    giПодРежимККМ = DrvFR.ECRAdvancedMode
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
    ' MsgBox "ошибка в процедуре StatusKKM"

End Sub
'+++++++++ Штрих-ФР-К v.
Public Sub CheckKKM(СheсkType As Byte, GAS As Double, Cost As Double, Npost As Byte)
    '        On Error GoTo err
    '        'Заносим в массив печати чеков информацию по всем постам
    '        With gGasStation(Npost)
    '            .gbСheсkType = СheсkType 'тип отбиваемого чека:0-после отпуска газа;1,2-не отбиваеться;3-"пустой" чек; 4-чек сторно
    '            .gdGasKKM = GAS 'CDbl(Format(Gas, "#####0.00")) 'количество газа, которое необходимо отразить в чеке
    '            .gdCostKKM = Cost 'цена при печати чека
    '            .gdFlagCheck = True 'поднимаем признак необходимости печати чека
    '        End With
    '        Exit Sub
    'err:
    '        MsgBox "ошибка в процедуре CheckKKM"
    '        End

End Sub
Public Sub PrintCheckKKM(ByVal NumPost As Byte)    'процедура вызывется раз в 3(три) секунды для проверки исправности ККМ и печати чеков

    '        Dim i As Integer
    '        Dim str As String
    '        Dim ModeKKM As Byte
    '        Dim ModeExKKM As Byte
    '        On Error GoTo err
    '        giErrorKKM = DrvFR.ResultCode
    '        gsErrorKKM = DrvFR.ResultCodeDescription
    '        StatusKKM
    '        If (giErrorKKM <> 0) And (giErrorKKM <> &H50) And (giErrorKKM <> &H8E) And (giErrorKKM <> &H1) And (giErrorKKM <> &H2) And (giErrorKKM <> &H6) Then 'Если произошедшая ошибка не является следующей:
    '        Else 'Если ошибок нет: ККМ работает нормально, то далее:
    '            ModeKKM = DrvFR.ECRMode
    '            ModeExKKM = DrvFR.ECRAdvancedMode
    '            If ((ModeKKM = 2) Or (ModeKKM = 4)) And (ModeExKKM = 0) Then
    '                If gGasStation(NumPost).gdFlagCheck = True Then 'Если поднят флаг на печать,то выполняем
    '                    'Оформление чека продажи
    '                    DrvFR.Password = pass
    '                    DrvFR.Quantity = gGasStation(NumPost).gdGasKKM ' * gGasStation(NumPost).gdCostKKM / gGasStation(NumPost).gdCostKKM
    '                    DrvFR.price = gGasStation(NumPost).gdCostKKM
    '                    DrvFR.Department = NumPost + 1
    '                    DrvFR.Tax1 = 1
    '                    DrvFR.Tax2 = 0
    '                    DrvFR.Tax3 = 0
    '                    DrvFR.Tax4 = 0
    '                    Select Case gGasStation(NumPost).gbСheсkType 'Тип отбиваемого чека
    '                    Case 0 'Чек отбиваеться после произведения отпуска газа населению
    '                      DrvFR.StringForPrinting = "Газ природный"
    '                      DrvFR.Sale
    '                    Case 1 'Чек не отбиваеться
    '                      Exit Sub
    '                    Case 2 'Чек не отбиваеться
    '                      Exit Sub
    '                    Case 4 '!?! Чек отбиваеться для операции сторнирования прошедшей заправки !?!
    '                    Case 3 'Чек отбивается в ручную (без отпуска газа) в случае ошибки оператора: при отпуске газа населению без отбития чека по окончании заправки
    '                    Case Else
    '                    End Select
    '                    'Закрытие чека с печатью(отрезка по умалчанию выставляется в таблице свойств ККМ)
    '                    DrvFR.Password = pass
    '                    DrvFR.CheckSubTotal 'Подводим итог чека, т.е. в Summ1 заносим все Sale
    '                    DrvFR.Password = pass
    '                    If (gGasStation(NumPost).glTypeEnd = 2) And (gGasStation(NumPost).glTypeZapr = 0) And (gGasStation(NumPost).glSumm1 > 0) Then 'выбрана заправка на сумму за наличку
    '                        DrvFR.Summ1 = gGasStation(NumPost).glSumm1
    '                    End If
    '                    DrvFR.Summ2 = 0
    '                    DrvFR.Summ3 = 0
    '                    DrvFR.Summ4 = 0
    '                    DrvFR.DiscountOnCheck = 0 'Скидок нет
    '                    DrvFR.Tax1 = 1
    '                    DrvFR.Tax2 = 0
    '                    DrvFR.Tax3 = 0
    '                    DrvFR.Tax4 = 0
    '                    DrvFR.StringForPrinting = "===================================="
    '                    DrvFR.CloseCheck
    '                    Sleep 100
    '                    giErrorKKM = DrvFR.ResultCode
    '                    gsErrorKKM = DrvFR.ResultCodeDescription
    '                    'Если процедура печати безошибочна, то снимаем флаг печати иначе оставляем флаг
    '                    If (giErrorKKM = 0) Then ''And glOperator <> 0 Then
    '                        gGasStation(NumPost).gdFlagCheck = False
    '                        Set ZaprDN = StatDB.OpenRecordset("zapr", dbOpenDynaset) 'gas,number,price,operator,typeZapr,typeFinish,date=Max(Now) where number = " & Number)
    '                        ZaprDN.FindLast "number=" & NumPost & " and TYPEZAPR=0"
    '                        ZaprDN.Edit
    '                        ZaprDN("DATE") = Now
    '                        ZaprDN("TYPEZAPR") = 5 'признак отбития чека
    '                        ZaprDN.Update
    '                    Else
    '                    '!!!надо бы проанализировать ошибку и принять соответствующие действия!!!
    '                    End If
    '
    '                    Exit Sub 'В любом случае выходим из процедуры печать, т.к. нельзя печатать вподряд
    '                End If
    '            Else
    '            End If
    '        End If
    '        Exit Sub
    'err:
    '        MsgBox "ошибка в процедуре PrintCheckKKM"

End Sub
'+++++++++ Штрих-ФР-Ф v.03
Public Sub ConnectKKM()
    On Error GoTo err
    'Регистрируем драйвер ФР, предварительно переписав его в папку c:\windiws\system
    '        Shell "regsvr32.exe /c /s " & Chr(34) '& "c:\windiws\drvfr.dll" & Chr(34)
    'Объявляем DrvFR как класс описанный в драйвере
    Set DrvFR = CreateObject("AddIn.Drvfr")
    'А теперь используем на полную катушку прописывая DrvFR.свойство
    DrvFR.GetActiveLD
    DrvFR.GetParamLD
    DrvFR.SetActiveLD
    StatusKKM
    If giErrorKKM <> 0 Then
    End If
    Exit Sub
err:
    MsgBox "нет драйвера KKM"

End Sub



