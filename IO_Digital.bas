Attribute VB_Name = "IO_Digital"
Option Explicit

Private Const gl����� = &H2C0
Private Const configCN1 = gl����� + &H3  '   707
Private Const configCN2 = gl����� + &H7  '   711

Global Const A0 = gl����� + &H0     '����� 704
Global Const B0 = gl����� + &H1     '705
Global Const C0 = gl����� + &H2     '706

Global Const A1 = gl����� + &H4     '����� 708
Global Const B1 = gl����� + &H5     '709
Global Const C1 = gl����� + &H6     '710

' The Driver functions
Private Declare Function DIO_DriverInit Lib "DIO.DLL" (wTotalBoards As Integer) As Integer
Declare Sub DIO_DriverClose Lib "DIO.DLL" ()

' The DIO functions
Private Declare Sub DIO_OutputByte Lib "DIO.DLL" _
        (ByVal address As Integer, ByVal dataout As Byte)
Private Declare Function DIO_InputByte Lib "DIO.DLL" _
        (ByVal address As Integer) As Integer


Public gn������(48) As Sensor    '��������� �������� �� ������ TB-24P � TB-16P8R

Private gn48DIO(5)   As Long    '��������� ��������� ����� PET-48DIO


Public Function Init_DIO_Driver() As String
    Dim i As Integer
    Dim msg As String
    i = DIO_DriverInit(1)
    Select Case i
        Case 0: msg = "NoError"
        Case 1: msg = "DriverOpenError"
        Case 2: msg = "DriverNoOpen"
        Case 3: msg = "GetDriverVersionError"
        Case 4: msg = "InstallIrqError"
        Case 5: msg = "ClearIntCountError"
        Case 6: msg = "GetIntCountError"
        Case 7: msg = "ResetError"
        Case 8: msg = "RemoveIrqError"
        Case 9: msg = "GetTotalBoardError"
        Case 10: msg = "CardNotFound"
        Case 11: msg = "GetConfigError"
        Case 12: msg = "ExceedBoardNumber"
    End Select

    If i <> 0 Then
        MsgBox msg, vbExclamation, "Driver DIO"
    End If

    DIO_OutputByte configCN1, &H8B    '������������� CN1 : A0 -output, B0 & C0 - input
    DIO_OutputByte configCN2, &H8B    '������������� CN2 : A1 -output, B1 & C1 - input
    
    ' ��������� ����
    ' TODO ����� ���������� �� ROff
    DIO_OutputByte A0, 0
    DIO_OutputByte A1, 0
End Function

Public Sub update_gn48DIO()
    gn48DIO(0) = CInt(DIO_InputByte(A0))
    gn48DIO(1) = Not (CInt(DIO_InputByte(B0)))
    gn48DIO(2) = Not (CInt(DIO_InputByte(C0)))

    gn48DIO(3) = CInt(DIO_InputByte(A1))
    gn48DIO(4) = Not (CInt(DIO_InputByte(B1)))
    gn48DIO(5) = Not (CInt(DIO_InputByte(C1)))
End Sub

Public Sub update_gn������()
    Dim p           As Integer
    Dim r           As Integer
    Dim i           As Integer
    Dim j           As Integer
    For i = 0 To 5
        p = gn48DIO(i)
        For j = 0 To 7
            r = p Mod 2
            If r = 0 Then
                gn������(8 * i + j).Data = 0
            Else
                gn������(8 * i + j).Data = 1
            End If
            p = Int(p / 2)
        Next j
    Next i
End Sub

' ���� ������� ����������, true ���� ������ ����������
Function isHandControl() As Boolean
    isHandControl = Not (CBool(gn������(15).Data))
End Function

' �1 ������? true ���� ������
Function k1_isOpen() As Boolean
    k1_isOpen = CBool(gn������(21).Data)
End Function

' �2 ������? true ���� ������
Function k2_isOpen() As Boolean
    k2_isOpen = CBool(gn������(16).Data)
End Function

' �3 ������? true ���� ������
Function k3_isOpen() As Boolean
    k3_isOpen = CBool(gn������(17).Data)
End Function

' �4 ������? true ���� ������
Function k4_isOpen() As Boolean
    k4_isOpen = CBool(gn������(18).Data)
End Function

' �5 ������? true ���� ������
Function k5_isOpen() As Boolean
    k5_isOpen = CBool(gn������(19).Data)
End Function

' �6 ������? true ���� ������
Function k6_isOpen() As Boolean
    k6_isOpen = CBool(gn������(20).Data)
End Function

' �7 ������? true ���� ������
Function k7_isOpen() As Boolean
    k7_isOpen = CBool(gn������(23).Data)
End Function

' ����� ��������� ���?
Function isClutchOn() As Boolean
    isClutchOn = CBool(gn������(36).Data)
End Function
'������� ������� � port 1
Public Sub ROn(port As Integer, n As Integer)
    Dim b As Byte
    ' ������� ���������
    b = getSoftPortState(port)
    ' ������� ��� (1 ��������� ������ ���� ��� ���� � ����� ������)
    ' ������� ��� (1 ��������� �� ����� ������)
    b = b Or n

    If (isDebug) Then
        Debug.Print "������ 1 � �����: " & port & " n: " & n
        gn48DIO(getIndexByPort(port)) = b
        Ron_debug port, n
    Else
        DIO_OutputByte port, b
    End If
    ' ��������� ����� ����� DIO
    ���������
End Sub

'������� ������� � port 0
Public Sub ROff(port As Integer, n As Integer)
    Dim b As Byte
    ' ������� ���������
    b = getSoftPortState(port)
    ' ������� � (1 ��������� ������ ���� ��� ���� � ����� ������)
    b = b And n

    If (isDebug) Then
        Debug.Print "������ 0 � �����: " & port & " n: " & n
        gn48DIO(getIndexByPort(port)) = b
        Roff_debug port, n
    Else
        DIO_OutputByte port, b
    End If
    ' ��������� ����� ����� DIO
    ���������
End Sub

' ���������� ��������� ����� �� ������ ���������� ������
Private Function getSoftPortState(port As Integer) As Byte
    Dim i As Integer: i = getIndexByPort(port)
    getSoftPortState = gn48DIO(i)
    'If isDebug Then
    '    Debug.Print "getSoftPortState", "port " & port, "return " & getSoftPortState
    'End If
End Function

Private Function getIndexByPort(port As Integer) As Integer
    Select Case port
        Case A0: getIndexByPort = 0
        Case B0: getIndexByPort = 1
        Case C0: getIndexByPort = 2
        Case A1: getIndexByPort = 3
        Case B1: getIndexByPort = 4
        Case C1: getIndexByPort = 5
        Case Else: err.Raise -1, , "������������ ����� �����: " & port
    End Select
End Function


Private Sub Ron_debug(port As Integer, n As Integer)
    If (port = A1 And n = 2) Then ' ���� ���
        ggACL8113(14) = 0.8
        Debug.Print "��� ����������"
    ElseIf (port = A1 And n = 4) Then ' ������� �1
        gn48DIO(getIndexByPort(C0)) = getSoftPortState(C0) Or 32
        Debug.Print "��1 ������"
    ElseIf (port = A1 And n = 6) Then ' ������� �1 � ���� ���
        ggACL8113(14) = 0.8
        Debug.Print "��� ����������"
        gn48DIO(getIndexByPort(C0)) = getSoftPortState(C0) Or 32
        Debug.Print "��1 ������"
    ElseIf (port = A1 And n = 8) Then
        gn48DIO(getIndexByPort(C0)) = getSoftPortState(C0) Or 1
        Debug.Print "��2 ������"
    ElseIf (port = A1 And n = 16) Then
        gn48DIO(getIndexByPort(C0)) = getSoftPortState(C0) Or 2
        Debug.Print "��3 ������"
    ElseIf (port = A1 And n = 24) Then ' ������� �2 � �3
        gn48DIO(getIndexByPort(C0)) = getSoftPortState(C0) Or 1
        Debug.Print "��2 ������"
        gn48DIO(getIndexByPort(C0)) = getSoftPortState(C0) Or 2
        Debug.Print "��3 ������"
    ElseIf (port = A1 And n = 32) Then
        gn48DIO(getIndexByPort(C0)) = getSoftPortState(C0) Or 4
        Debug.Print "��4 ������"
    ElseIf (port = A1 And n = 34) Then '���� ���, ������� ���4
        ggACL8113(14) = 0.8
        Debug.Print "��� ����������"
        gn48DIO(getIndexByPort(C0)) = getSoftPortState(C0) Or 4
        Debug.Print "��4 ������"
    ElseIf (port = A1 And n = 42) Then ' C��� ���, ������� ���2, ������� ���4
        ggACL8113(14) = 0.8
        Debug.Print "��� ����������"
        gn48DIO(getIndexByPort(C0)) = getSoftPortState(C0) Or 1
        Debug.Print "��2 ������"
        gn48DIO(getIndexByPort(C0)) = getSoftPortState(C0) Or 4
        Debug.Print "��4 ������"
    ElseIf (port = A1 And n = 64) Then
        gn48DIO(getIndexByPort(C0)) = getSoftPortState(C0) Or 8
        Debug.Print "��5 ������"
    ElseIf (port = A1 And n = 128) Then
        gn48DIO(getIndexByPort(C0)) = getSoftPortState(C0) Or 16
        Debug.Print "��6 ������"
    ElseIf (port = A0 And n = 2) Then
        gn48DIO(getIndexByPort(C0)) = getSoftPortState(C0) Or 128
        Debug.Print "��7 ������"
    ElseIf (port = B0 And n = 128) Then
        Debug.Print "����. ������ ����������"
    ElseIf (port = B1 And n = 16) Then
        Debug.Print "��� ����� ���������"
    Else
        Debug.Print "�������������� ������� �������", port, n
    End If
End Sub

Private Sub Roff_debug(port As Integer, n As Integer)
    If (port = A1 And n = 0) Then
        gn48DIO(getIndexByPort(C0)) = 0
        Debug.Print "��� �� �������"
    ElseIf (port = A1 And n = 1) Then ' ������� ��� ��, ��� ���� 2
        gn48DIO(getIndexByPort(C0)) = 0
        Debug.Print "��� �� �������, ���� 2 ���"
    ElseIf (port = A1 And n = 239) Then ' ������� �3
        gn48DIO(getIndexByPort(C0)) = getSoftPortState(C0) And Not (2)
        Debug.Print "�3 ������"
    ElseIf (port = A1 And n = 223) Then ' ������� �4
        gn48DIO(getIndexByPort(C0)) = getSoftPortState(C0) And Not (4)
        Debug.Print "�4 ������"
    ElseIf (port = A1 And n = 127) Then ' ������� �6
        gn48DIO(getIndexByPort(C0)) = getSoftPortState(C0) And Not (16)
        Debug.Print "�6 ������"
    ElseIf (port = A1 And n = 191) Then ' ������� �5
        gn48DIO(getIndexByPort(C0)) = getSoftPortState(C0) And Not (8)
        Debug.Print "�5 ������"
    ElseIf (port = B1 And n = Not (16)) Then
        Debug.Print "���� ����� ���������"
    Else
        Debug.Print "�������������� ������� �������", port, n
    End If
End Sub



' TODO ��������� ������������
' ROff A1, 1 �� ������ ��������� �1-6, �� � ��������� ���� "���� ����"