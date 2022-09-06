Attribute VB_Name = "IO_File"
Option Explicit

Private rec As pswd
Private Password As String

'структура для secret file (gdK_file_name)
Public Type pswd
    PC              As Double
    pwd             As String * 7
End Type

Private Function gdK_file_name() As String
   gdK_file_name = App.Path + "\agnks.config"
End Function

'Функция инициализации данных, считывание с диска
'TODO Вывести на отдельную вкладку информацию
' о моточасах
' поправочный коэфф
' цена газа
' плотность газа
' смена пароля и настройка параметров прямо из окна пульта
Public Function InitDisk() As Integer
    'Возвращает код ошибки : 0 -нет ошибок
   ' On Error Resume Next
   'TODO Добавить смену курсора внутрь функций
    frmStart.MousePointer = vbHourglass

    init_Database
    'Получаем моточасы из базы
    GMC = getGMC_from_DB

    load_statistic_from_DB 'TODO вынести из функции инициализации диска?

    init_gdK_file
    init_price_file
    init_data_file
    frmStart.MousePointer = vbArrow
    InitDisk = 0
    Exit Function

ErrorHandler:        'Если есть какие-нибудь ошибки возвращаем -1
    InitDisk = -1
    Exit Function
End Function

Private Sub init_gdK_file()
    Dim fh As Long: fh = FreeFile
    Dim fLen As Long
    On Error Resume Next
        fLen = FileLen(gdK_file_name)
    On Error GoTo 0
    If fLen = 0 Then
        Password = "LAB"
        gdK = 0
        MsgBox "Отсутвует файл конфигурации АГНКС: " & gdK_file_name, vbExclamation
    Else
        Open gdK_file_name For Random As fh Len = Len(rec)
            Get #fh, 1, rec
            Password = Trim(rec.pwd)
            gdK = rec.PC
        Close #fh
    End If
    frmStart.lblPC.Caption = Format(gdK, "0.000")
End Sub

'FIXME обработать отмену ввода
Sub setting_gdK()
    Dim fh As Long
    Dim s As String
    Dim title As String: title = "DANGER - Обновление поправочного коэффициента"
    s = InputBox("Введите пароль", title)
    If (s = Password) Then
        s = InputBox("Введите поправочный коэффициент", title, Format(gdK, "0.000"))
        If (CDbl(s) > 0) And (CDbl(s) <= 10) Then
            gdK = CDbl(s)
            fh = FreeFile
            Open gdK_file_name For Random As fh Len = Len(rec)
                rec.pwd = Password
                rec.PC = gdK
                Put #fh, 1, rec
            Close #fh
            MsgBox "Коэффициент введен", vbInformation
        End If
    Else
        MsgBox "Пароль не верный", vbCritical
    End If
    init_gdK_file
End Sub

Sub update_gdK_pass()
    Dim fh As Long
    Dim s As String
    Dim s1 As String
    Dim title As String: title = "DANGER - Обновление пароля"
    s = InputBox("Введите текущий пароль", title)
    If (s = Password) Then
        s = InputBox("Введите новый пароль", title)
        If (Len(s) > 0) And (Len(s) <= 7) Then
            s1 = InputBox("Повторите новый пароль", title)
            If (s = s1) Then
                Password = s1
                fh = FreeFile
                Open gdK_file_name For Random As fh Len = Len(rec)
                    rec.pwd = Password
                    rec.PC = gdK
                    Put #fh, 1, rec
                Close #fh
                MsgBox "Пароль введен", vbInformation
            Else
                MsgBox "Пароли не совпадают", vbCritical
            End If
        End If
    Else
        MsgBox "Пароль не верный", vbCritical
    End If
    init_gdK_file
End Sub

Private Sub init_price_file()
    Dim fh As Long: fh = FreeFile
    Dim s As String
    s = App.Path & "\price.txt"
    
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
        Line Input #fh, s
        gdPrice = CDbl(s)
        
        Line Input #fh, s
        gdPlot = CDbl(s)
    Close #fh

    frmStart.Label_Price.Caption = gdPrice
    If gdPlot < 0.5 Or gdPlot > 1 Then gdPlot = 0.7
End Sub

Private Sub init_data_file()
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
