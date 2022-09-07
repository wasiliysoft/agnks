Attribute VB_Name = "IO_File"
Option Explicit

Private pAgnksÑonfig As AgnksÑonfigType

'ñòðóêòóðà äëÿ êîíôèãóðàöèîííîãî ôàéëà (configFilePath)
Public Type AgnksÑonfigType
    PC As Double        ' Ïîïðàâî÷íûé êîýôôèöèåíò
    price As Double     ' Öåíà ãàçà
    plot As Double      ' Ïëîòíîñòü ãàçà
    motoMinute As Long  ' Ìîòîðåñóðñ â ìèíóòàõ
    pwd As String * 7   ' Ïàðîëü
End Type

Private Function configFilePath() As String
   configFilePath = App.Path + "\agnks.config"
End Function

Function agnksÑonfig() As AgnksÑonfigType
   agnksÑonfig = pAgnksÑonfig
End Function


'Ôóíêöèÿ èíèöèàëèçàöèè äàííûõ, ñ÷èòûâàíèå ñ äèñêà
'TODO Âûâåñòè íà îòäåëüíóþ âêëàäêó èíôîðìàöèþ
' î ìîòî÷àñàõ
' ïîïðàâî÷íûé êîýôô
' öåíà ãàçà
' ïëîòíîñòü ãàçà
' ñìåíà ïàðîëÿ è íàñòðîéêà ïàðàìåòðîâ ïðÿìî èç îêíà ïóëüòà
Public Function InitDisk() As Integer
    'Âîçâðàùàåò êîä îøèáêè : 0 -íåò îøèáîê
   ' On Error Resume Next
   'TODO Äîáàâèòü ñìåíó êóðñîðà âíóòðü ôóíêöèé
    frmStart.MousePointer = vbHourglass

    init_Database
    'Ïîëó÷àåì ìîòî÷àñû èç áàçû
    GMC = getGMC_from_DB

    load_statistic_from_DB 'TODO âûíåñòè èç ôóíêöèè èíèöèàëèçàöèè äèñêà?

    init_agnksConfig
    
    init_SensorDescr_file
    frmStart.MousePointer = vbArrow
    InitDisk = 0
    Exit Function

ErrorHandler:        'Åñëè åñòü êàêèå-íèáóäü îøèáêè âîçâðàùàåì -1
    InitDisk = -1
    Exit Function
End Function

Private Sub init_agnksConfig()
    Dim fh As Long: fh = FreeFile
    Dim fLen As Long
    On Error Resume Next
        fLen = FileLen(configFilePath)
    On Error GoTo 0
    If fLen = 0 Then
        With pAgnksÑonfig
            .motoMinute = -1
            .PC = -1
            .plot = -1
            .price = -1
            .pwd = "LAB"
        End With
        MsgBox "Îòñóòâóåò ôàéë êîíôèãóðàöèè ÀÃÍÊÑ: " & configFilePath, vbExclamation
    Else
        Open configFilePath For Random As fh Len = Len(pAgnksÑonfig)
            Get #fh, 1, pAgnksÑonfig
        Close #fh
    End If
    frmStart.lblPC.Caption = Format(agnksÑonfig.PC, "0.000")
    frmStart.Label_Price.Caption = pAgnksÑonfig.price
    frmStart.lbl_gnPlot.Caption = pAgnksÑonfig.plot
End Sub

'FIXME îáðàáîòàòü îòìåíó ââîäà
Sub updatePC()
    Dim s As String
    Dim title As String: title = "DANGER - Îáíîâëåíèå ïîïðàâî÷íîãî êîýôôèöèåíòà"
    s = InputBox("Ââåäèòå ïàðîëü", title)
    If (s = Trim(pAgnksÑonfig.pwd)) Then
        s = InputBox("Ââåäèòå ïîïðàâî÷íûé êîýôôèöèåíò", title, Format(agnksÑonfig.PC, "0.000"))
        If (CDbl(s) > 0) And (CDbl(s) <= 10) Then
            pAgnksÑonfig.PC = CDbl(s)
            saveConfig
            MsgBox "Êîýôôèöèåíò ââåäåí", vbInformation
        End If
    Else
        MsgBox "Ïàðîëü íå âåðíûé", vbCritical
    End If
    init_agnksConfig
End Sub

' TODO implementation
Sub updatePlot()
    Dim d As Double: d = 0
    Dim sInput As String 
    sInput = InputBox("Ââåäèòå íîâîå çíà÷åíèå ïëîòíîñòè ãàçà", , Format(agnksÑonfig.plot, "0.0000"))
    If (Len(sInput) = 0) Then Exit Sub
    On Error Resume Next
        d = CDbl(sInput)
    On Error GoTo 0 
    If  d >= 1 Or d <= 0.5 Then
        Msgbox "Íåêîððåêòíûé ââîä", vbExclamation
    Else
        pAgnksÑonfig.plot = CDbl(d)
        saveConfig
        init_agnksConfig
        MsgBox "Îáíîâëåíî", vbInformation
    End If
End Sub

' TODO implementation
Sub updatePrice()
    pAgnksÑonfig.price = CDbl(11.7)
    saveConfig
    init_agnksConfig
End Sub
' TODO return result
Private Sub saveConfig()
    Dim fh As Long: fh = FreeFile
    Open configFilePath For Random As fh Len = Len(pAgnksÑonfig)
        Put #fh, 1, pAgnksÑonfig
    Close #fh
End Sub

Sub updatePWD()
    Dim s As String
    Dim s1 As String
    Dim title As String: title = "DANGER - Îáíîâëåíèå ïàðîëÿ"
    s = InputBox("Ââåäèòå òåêóùèé ïàðîëü", title)
    If (s = Trim(pAgnksÑonfig.pwd)) Then
        s = InputBox("Ââåäèòå íîâûé ïàðîëü", title)
        If (Len(s) > 0) And (Len(s) <= 7) Then
            s1 = InputBox("Ïîâòîðèòå íîâûé ïàðîëü", title)
            If (s = s1) Then
                pAgnksÑonfig.pwd = s1
                saveConfig
                MsgBox "Ïàðîëü ââåäåí", vbInformation
            Else
                MsgBox "Ïàðîëè íå ñîâïàäàþò", vbCritical
            End If
        End If
    Else
        MsgBox "Ïàðîëü íå âåðíûé", vbCritical
    End If
    init_agnksConfig
End Sub

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
        MsgBox "Îòñóòâóåò ôàéë êîíôèãóðàöèè ÀÃÍÊÑ: " & s, vbExclamation
        Exit Sub
    End If

    Open s For Input Access Read As fh
        Seek #fh, 1
        'Ââîä ïîÿñíåíèé î äàò÷èêàõ DIO
        For i = 0 To 47
            Line Input #fh, gnÄàò÷èê(i).Note
            frmStart.Label2(i).Caption = gnÄàò÷èê(i).Note
        Next i
        'Ââîä ïîÿñíåíèé î äàò÷èêàõ ISO
        For i = 0 To 15
            Line Input #fh, s
            frmStart.Text2(i).Text = s
        Next i
    Close #fh
End Sub
