Attribute VB_Name = "IO_Analog"
Option Explicit

Public Function Init_ISO813_Driver() As String

    Dim i           As Integer
    glaАдрес = Val("&H220")
    glРезультат = ISO813_DriverInit()
    If glРезультат <> ISO813_NoError Then
        i = MsgBox("Can not initial Driver!!!", , "ISO813 Card Error")
    ElseIf glРезультат = 2 Then
        Init_ISO813_Driver = "Driver open error !"
    Else
        Init_ISO813_Driver = "Плата ACL8113 в норме"
    End If
End Function
