Attribute VB_Name = "IO_Analog"
Option Explicit

Public Function Init_ISO813_Driver() As String

    Dim i           As Integer
    gla����� = Val("&H220")
    gl��������� = ISO813_DriverInit()
    If gl��������� <> ISO813_NoError Then
        i = MsgBox("Can not initial Driver!!!", , "ISO813 Card Error")
    ElseIf gl��������� = 2 Then
        Init_ISO813_Driver = "Driver open error !"
    Else
        Init_ISO813_Driver = "����� ACL8113 � �����"
    End If
End Function
