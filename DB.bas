Attribute VB_Name = "DB"
Option Explicit


Public StatDB       As Database
Public SelectRS     As Recordset
Private StatRS       As Recordset
Private StatWS       As Workspace


Sub init_Database()
    Dim s As String
    s = App.Path + "\base.mdb"
    Set StatWS = DBEngine.Workspaces(0)
    Set StatDB = StatWS.OpenDatabase(s)
    Set StatRS = StatDB.OpenRecordset("stat", dbOpenTable)
    'TODO создать базу если пустая
End Sub

Function getGMC_from_DB() As Long
    Set SelectRS = StatDB.OpenRecordset("SELECT * From stat ORDER BY stat.data DESC")
    If Not IsNull(SelectRS(0)) Then
        getGMC_from_DB = SelectRS("MOTO")
    Else
        getGMC_from_DB = 0
    End If
    SelectRS.Close
    Set SelectRS = Nothing
End Function

Sub saveGMC_in_DB()
    Set SelectRS = StatDB.OpenRecordset("SELECT * From stat ORDER BY stat.data DESC")
    SelectRS.Edit
    GMC = GMC + MotorCount
    MotorCount = 0
    SelectRS("MOTO") = GMC
    SelectRS.Update
    SelectRS.Close
    Set SelectRS = Nothing
End Sub
Sub load_statistic_from_DB()
'
' Заполняет вкладку "Жарнал" данными из базы.
'
    Dim i           As Integer

    Dim dateMin     As Date
    Dim dateMax     As Date
    Dim dateNow     As Date: dateNow = Now

    Dim s           As String
    Dim s1          As String

   frmStart.lblStat(1).Caption = "За " + Format(Now, "mmmm")
   frmStart.lblStat(2).Caption = "За " + Format(Now, "yyyy") + " год"

    If StatRS.EOF Then
        'пустая база данных
        Exit Sub
    End If

    'Получаем самую раннюю дату
    Set SelectRS = StatDB.OpenRecordset("select MIN(DATA) from stat ")
    dateMin = SelectRS(0)
    'Получаем самую позднюю дату
    Set SelectRS = StatDB.OpenRecordset("select MAX(DATA) from stat ")
    dateMax = SelectRS(0)
    
    'Для заполнения 4 колонки в цикле посчитать все суммарные заправки
    'за годы начиная с самого раннего и до предыдущего
    frmStart.lstStat(3).Clear
    For i = Year(dateMin) To Year(dateNow) - 1
        s = "#1/1/" + CStr(i) + " 00:00:00#"
        s1 = "#12/31/" + CStr(i) + " 23:59:59#"
        Set SelectRS = StatDB.OpenRecordset("select SUM(GAZ_CAR) from stat where (stat.DATA between " & s & " AND " & s1 & ")")
        If Not IsNull(SelectRS(0)) Then
            s = Format(CStr(i)) + "        " + Format(CStr(SelectRS(0)), "###0.00")
            frmStart.lstStat(3).AddItem (s)
        End If
    Next i

    'Колонка "За год" в цикле посчитать все суммарные заправки
    'за месяцы этого года начиная с января и до предыдущего месяца.
    frmStart.lstStat(2).Clear
    If Year(dateMax) = Year(dateNow) Then
        For i = 1 To Month(dateNow) - 1
            s = "#" & i & "/1/" & Year(dateNow) & " 00:00:00#"
            s1 = "#" & i & "/" & lastDayByMonth(i, Year(dateNow)) & "/" & Year(dateNow) & " 23:59:59#"
            Set SelectRS = StatDB.OpenRecordset("select SUM(GAZ_CAR) from stat where (stat.DATA between " & s & " AND " & s1 & ")")
            If Not IsNull(SelectRS(0)) Then
                s = Format(i, "00") + "        " + Format(SelectRS(0), "###0.00")
                frmStart.lstStat(2).AddItem (s)
            End If
        Next i
    End If

    'Колонка "За месяц" в цикле посчитать все суммарные заправки за дни начиная с 1 этого месяца и до пред.
    frmStart.lstStat(1).Clear
    If Month(dateMax) = Month(dateNow) Then
        For i = 1 To Day(dateNow) - 1
            s = "#" & Month(dateNow) & "/" & i & "/" & Year(dateNow) & " 00:00:00#"
            s1 = "#" & Month(dateNow) & "/" & (i + 1) & "/" & Year(dateNow) & " 23:59:59#"
            Set SelectRS = StatDB.OpenRecordset("select SUM(GAZ_CAR) from stat where (stat.DATA between " & s & " AND " & s1 & ")")
            If Not IsNull(SelectRS(0)) Then
                s = Format(i, "00") + "        " + Format(SelectRS(0), "###0.00")
                frmStart.lstStat(1).AddItem (s)
            End If
        Next i
    End If

    'Колонка "За сегодня"
    update_Journal_col_1
End Sub

Private Sub update_Journal_col_1()
    Dim s As String, s1 As String
    Dim i As Long
    frmStart.lstStat(0).Clear
    s = Format(Now, "\#mm\/dd\/yyyy 00:00:00\#")
    s1 = Format(Now, "\#mm\/dd\/yyyy 23:59:59\#")
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
End Sub

Sub StatRS_Insert()
    Dim v As Double: v = gdРасход1 / gdPlot
    If v > 0.1 Then ' Защита от околонулевых записей заправок
        GMC = GMC + MotorCount
        MotorCount = 0
        StatRS.AddNew
        StatRS("DATA") = Now
        StatRS("GAZ_CAR") = gdРасход1 / gdPlot        '* 1.42
        StatRS("GAZ_IR1") = gdИР1
        StatRS("MOTO") = GMC
        StatRS.Update
        update_Journal_col_1
    End If
    ' FIXME после добавления в базу нужно обновить
    ' колонки "За месяц", "За год"
    ' при этом учитывать возможность работы 24/7
    ' если произойдет смена даты или месяца
    ' то и колонки должны обновится
End Sub

Public Function lastDayByMonth(ByVal m, ByVal yyyy) As Integer
    Select Case m
        Case 1, 3, 5, 7, 8, 10, 12
            lastDayByMonth = 31
        Case 2
            Select Case yyyy
                Case 2004, 2008, 2012, 2016, 2020, 2024, 2028, 2032, 2036, 2040
                    lastDayByMonth = 29
                Case Else
                    lastDayByMonth = 28
            End Select
        Case Else
            lastDayByMonth = 30
    End Select
End Function

'Public Type StatRowType
'    dt              As Date
'    IR1             As Double
'    IR2             As Double
'    Motor           As Long
'End Type
