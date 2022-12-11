Attribute VB_Name = "IO_File"
Option Explicit

Private pAgnksConfig As AgnksConfigType

'структура для конфигурационного файла (configFilePath)
Public Type AgnksConfigType
    PC As Double        ' Поправочный коэффициент
    Price As Double     ' Цена газа
    plot As Double      ' Плотность газа
    pwd As String * 10   ' Пароль
End Type

Private Function configFilePath() As String
   configFilePath = App.Path + "\agnks.config"
End Function

Function agnksConfig() As AgnksConfigType
   agnksConfig = pAgnksConfig
End Function

Public Sub InitDisk()
   ' On Error Resume Next
   'TODO Добавить смену курсора внутрь функций работы с базой
    frmStart.MousePointer = vbHourglass

    init_agnksConfig
    init_SensorDescr_file
    init_Database
    load_statistic_from_DB 'TODO вынести из функции инициализации диска?
    'Получаем моточасы из базы
    GMC = getGMC_from_DB

    frmStart.MousePointer = vbArrow
End Sub

Private Sub init_agnksConfig()
    Dim fh As Long: fh = FreeFile
    Dim fLen As Long
    On Error Resume Next
        fLen = FileLen(configFilePath)
    On Error GoTo 0
    If fLen = 0 Then
        With pAgnksConfig
            .PC = -1
            .plot = -1
            .Price = -1
            .pwd = "LAB"
        End With
        MsgBox "Отсутвует файл конфигурации АГНКС: " & configFilePath, vbExclamation
    Else
        Open configFilePath For Random As fh Len = Len(pAgnksConfig)
            Get #fh, 1, pAgnksConfig
        Close #fh
    End If
    frmStart.lblPC.Caption = Format(agnksConfig.PC, "0.0000")
    frmStart.Price.Caption = Format(agnksConfig.Price, "0.00")
    frmStart.lbl_gnPlot.Caption = Format(pAgnksConfig.plot, "0.0000")
End Sub


Sub updatePC()
    If isAuth = False Then Exit Sub

    Dim d As Double: d = 0
    Dim sInput As String
    sInput = InputBox("Введите поправочный коэффициент", , Format(agnksConfig.PC, "0.0000"))
    If (Len(sInput) = 0) Then Exit Sub
    On Error Resume Next
        d = CDbl(sInput)
    On Error GoTo 0

    If d < -10 Or 10 < d Or d <> Round(d, 4) Then
        MsgBox "Разрешен ввод от -10 до 10 с точностью 4 знака после запятой.", vbExclamation, "Некорректный ввод"
    Else
        pAgnksConfig.PC = d
        saveConfig
        init_agnksConfig
        MsgBox "Обновлено", vbInformation
    End If
End Sub


Sub updatePlot()
    If isAuth = False Then Exit Sub

    Dim d As Double: d = 0
    Dim sInput As String
    sInput = InputBox("Введите новое значение плотности газа", , Format(agnksConfig.plot, "0.0000"))
    If (Len(sInput) = 0) Then Exit Sub
    On Error Resume Next
        d = CDbl(sInput)
    On Error GoTo 0
    If d >= 1 Or d <= 0.5 Or d <> Round(d, 4) Then
        MsgBox "Допустимое значение от 0,5 до 1 с точностью 4 знака после запятой.", vbExclamation, "Некорректный ввод"
    Else
        pAgnksConfig.plot = d
        saveConfig
        init_agnksConfig
        MsgBox "Обновлено", vbInformation
    End If
End Sub


Sub updatePrice()
    If isAuth = False Then Exit Sub

    Dim d As Double: d = 0
    Dim sInput As String
    sInput = InputBox("Введите новое значение цены газа", , Format(agnksConfig.Price, "0.00"))
    If (Len(sInput) = 0) Then Exit Sub
    On Error Resume Next
        d = CDbl(sInput)
    On Error GoTo 0
    If d >= 1000 Or d <= 0 Or d <> Round(d, 2) Then
        MsgBox "Допустимое значение от 0 до 1000 с точностью 2 знака после запятой.", vbExclamation, "Некорректный ввод"
    Else
        pAgnksConfig.Price = d
        saveConfig
        init_agnksConfig
        MsgBox "Обновлено", vbInformation
    End If
End Sub

Sub updatePWD()
    If isAuth = False Then Exit Sub

    Dim sInput1 As String
    Dim sInput2 As String
    
    ' ВНИМАНИЕ!!!
    ' Максимальная длинна хранимого пароля
    ' определяется в структуре AgnksConfigType

    frmPassword.lblDescription = "Введите новый пароль, от 3 до 10 символов"
    frmPassword.txtPassword = ""
        frmPassword.Show vbModal
        sInput1 = frmPassword.txtPassword
    frmPassword.txtPassword = ""

    sInput1 = Trim(sInput1)
    If Len(sInput1) = 0 Then ' Отмена ввода
        Exit Sub
    ElseIf Len(sInput1) < 3 Or 10 < Len(sInput1) Then
        MsgBox "Некорректный пароль", vbExclamation
        Exit Sub
    Else
        frmPassword.lblDescription = "Повторите новый пароль"
        frmPassword.txtPassword = ""
            frmPassword.Show vbModal
            sInput2 = frmPassword.txtPassword
        frmPassword.txtPassword = ""

        sInput2 = Trim(sInput2)
        If sInput1 <> sInput2 Then
            MsgBox "Пароли не совпадают", vbExclamation
            Exit Sub
        Else
            pAgnksConfig.pwd = sInput2
            saveConfig
            init_agnksConfig
            MsgBox "Обновлено", vbInformation
        End If
    End If
End Sub

' TODO return result
Private Sub saveConfig()
    Dim fh As Long: fh = FreeFile
    Open configFilePath For Random As fh Len = Len(pAgnksConfig)
        Put #fh, 1, pAgnksConfig
    Close #fh
End Sub

Function isAuth() As Boolean
    isAuth = False
    Dim sInput As String
    frmPassword.lblDescription = "Введите пароль"
    frmPassword.txtPassword = ""
        frmPassword.Show vbModal
        sInput = frmPassword.txtPassword
    frmPassword.txtPassword = ""

    sInput = Trim(sInput)
    If Len(sInput) = 0 Then
        Exit Function
    ElseIf sInput <> Trim(pAgnksConfig.pwd) Then
        MsgBox "Неверный пароль", vbExclamation
        Exit Function
    Else
        isAuth = True
    End If
End Function

Private Sub init_SensorDescr_file()
    Dim fh As Long: fh = FreeFile
    Dim s As String
    Dim i As Integer
    
    s = App.Path & "\data.txt"
    Dim fLen As Long
    On Error Resume Next
        fLen = FileLen(s)
    On Error GoTo 0
    If fLen = 0 Then
        MsgBox "Отсутвует файл конфигурации АГНКС: " & s, vbExclamation
        Exit Sub
    End If

    Open s For Input Access Read As fh
        Seek #fh, 1
        'Ввод пояснений о датчиках DIO
        For i = 0 To 47
            Line Input #fh, gnДатчик(i).Note
            frmStart.Label2(i).Caption = gnДатчик(i).Note
        Next i
        'Ввод пояснений о датчиках ISO
        For i = 0 To 15
            Line Input #fh, s
            frmStart.Text2(i).Text = s
        Next i
    Close #fh
End Sub
