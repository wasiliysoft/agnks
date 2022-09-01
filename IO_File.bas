Attribute VB_Name = "IO_File"
Option Explicit

Private Const gdK_file_name = "\agnks.config"
Private rec As pswd
Private Password As String

'структура для secret file (gdK_file_name)
Public Type pswd
    PC              As Double
    pwd             As String * 7
End Type

'Функция инициализации данных, считывание с диска
'TODO Вывести на отдельную вкладку информацию
' о моточасах
' поправочный коэфф
' цена газа
' плотность газа
' смена пароля и настройка параметров прямо из окна пульта
Public Function InitDisk() As Integer
    'Возвращает код ошибки : 0 -нет ошибок
    Dim i           As Integer
    Dim j           As Integer
    Dim j1          As Integer

    Dim k           As Integer
    Dim t           As MyRecType
    Dim temp1       As MyRecType
    Dim temp2       As MyRecType
    Dim d           As Date
    Dim sum         As Double
    Dim idx1, idx2, idx3 As Integer
    'Для ввода поправочного коэффициента
    Dim descr       As Integer
    Dim s           As String
    Dim s1          As String
    Dim v           As Variant
   ' On Error Resume Next
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
        gDateRec = Now    ' temp2.dt

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
    Else
        GMC = 0
        gDateRec = Now    ' temp2.dt

    End If

    gDateRec = Now
    init_gdK_file
    init_price_file
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
End Sub

'FIXME обработать отмену ввода
Sub setting_gdK()
    Dim fh As Long
    Dim s As String
    s = InputBox("Введите пароль", "DANGER")
    If (s = Password) Then
        s = InputBox("Введите поправочный коэффициент", "DANGER", Format(gdK, "0.000"))
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
End Sub

Sub update_gdK_pass()
    Dim fh As Long
    Dim s As String
    Dim s1 As String
    s = InputBox("Введите пароль", "DANGER")
    If (s = Password) Then
        s = InputBox("Введите новый пароль", "DANGER")
        If (Len(s) > 0) And (Len(s) <= 7) Then
            s1 = InputBox("Повторите новый пароль", "DANGER")
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
End Sub
Private Sub init_price_file()
    Dim fh As Long: fh = FreeFile
    Dim s As String
    
    s = App.Path & "\price.txt"
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
