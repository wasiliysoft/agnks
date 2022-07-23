Attribute VB_Name = "Module2"

Public Function Инициал_ACL8113() As String

    Dim i As Integer
        glaАдрес = Val("&H220")
        glНомерПлаты = 0
       ' glРезультат = W_8113_Initial(glНомерПлаты, glАдрес)
    glРезультат = ISO813_DriverInit()
    If glРезультат <> ISO813_NoError Then
        i = MsgBox("Can not initial Driver!!!", , "ISO813 Card Error")
    ElseIf glРезультат = 2 Then
           Инициал_ACL8113 = "Driver open error !"
    Else
    
       Инициал_ACL8113 = "Плата ACL8113 в норме"
    End If

    

End Function



Public Function Инициал_Pet48DIO() As String
    Dim Dummy
    Dim wTotalBoards As Integer
    wTotalBoards = 1
    'Инициализация
    
        glАдрес = Val("&H2C0") 'Оставляю по умолчанию
       ' glНомерПлаты = 0
        glIRQ = 3
       ' glРезультат = W_48DIO_Initial(glНомерПлаты, glАдрес, glIRQ)
       glРезультат = DIO_DriverInit(wTotalBoards)
  
    If glРезультат <> DIO_NoError Then
        MsgBox "Driver DIO Initialize OK!!"
    Else
        Инициал_Pet48DIO = "Плата Pet48DIO в норме"
        ' Don't forget to close the driver by DIO_DriverClose()
    End If
       DIO_OutputByte glАдрес + &H3, &H8B 'Устанавливаем CN1 : A0 -output, B0 & C0 - input
       DIO_OutputByte glАдрес + &H7, &H8B 'Устанавливаем CN2 : A1 -output, B1 & C1 - input
       
       DIO_OutputByte glАдрес, 0
       DIO_OutputByte glАдрес + &H4, 0
       
    'Выключить реле 0 (порт A1)
       ' glРезультат = W_48DIO_DO(256, 0)
        
    'Выключить реле 0 (порт A0)
       ' glРезультат = W_48DIO_DO(0, 0)
    
    
End Function


Public Function ИниКонтроль() As String
    ИниКонтроль = ""
'Инициализация платы ACL8113
   gsРезультат = Инициал_ACL8113
'Инициализация платы Pet48DIO
    gsРезультат = Инициал_Pet48DIO
    
    ИниКонтроль = "OK"
End Function


