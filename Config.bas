Attribute VB_Name = "Config"
'��� ���������� ���������� ��������� �������
Option Explicit

Global Const isDebug = true

Global Const gdUpLevel = 200 * 0.0981    '������ �������� ��� ��������

Global Const gdRashAkkEnd = 65    '������ ������ �� �������� ���������� ����� �� �������������

'����������� ��������
'Global Const ggMinPress = 197    ' ����������� �������� � �������������

Global Const conHwndTopmost = -1
Global Const conSwpNoActivate = &H10
Global Const conSwpShowWindow = &H40

Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long


'���������� ����������������� ������
'������ ������� + ��������
Public Type Sensor
    Data            As Integer
    Note            As String
End Type

'��� ������� ����
Public gd��1        As Double
Public gd��2        As Double


'��� ��������
Public gbOnlyAkk    As Boolean    '���� �������� ������ �� �������������
'������ ����� �� �������� ���
 '����� ��������
 ' 0 - �������� ���������
 ' 1 - ��������
 ' 2 - ��������
 ' 3 - Danger
Public giStage      As Integer   

'����� ���������
' 0 - 
' 1 - 
Public giStage1     As Integer

' 0 - 
' 1 - 
' 2 - ���� �������� �������� �� ��������� �5 � ������� �� ���� 3
' 3 - ���� � ��� ��������� �������� �� ������ �6, ������� �� ���� 4
' 4 - �������� ��� � ����
' 5 - 
' 6 - 
' 7 - �� ����� �������� ��� ������� � �������� ���� (�������� �5, ������� �� ���� 4)
' 8 - ����� ��������� ������ �� ��� (��������� �5 � �6, ������� �� ���� 9)
' 9 - ������� ������ �� ��� (������� ������� ���������� ������� ����� ��2.1 � ��2.2)
Public giStage2     As Integer    '�������� ����� ��������

Public gbFrmShow    As Boolean

' gbCmdStart = true
' giStage = 3,0
' gbCmdStart = false
' giStage = 1
Public gbCmdStart   As Boolean    '��� ������ ����\��������
Public gbAkkum      As Boolean
Public giTrigger    As Integer    ' ��� ����� �������
Public Car          As Integer
Public gb��������   As Boolean    '���������� ����� ����� � ������� ����� ������� ������
'���� �� �������� ������ ����������
Public gd������1    As Double    '������ �� �������� ����������
Public giMain������ As Integer    '���� ��� �������� ������� ����� ����:
' 1 - ���������
' 0 - ������
'-1 - ��������

Public gdK          As Double    '����������� �����������

' ���������� ��� ����������
Public glAver       As Long    ' ������ ����� ��� ����������
Public glCounter    As Long
Public sum(31)      As Double


Public gbRunDVS     As Boolean





Public FileHandle   As Integer    '������������� ����� � ����������
Public MotorCount   As Long    '������� ������������
Public GMC          As Long    '���������� �������


Public giCountZ     As Integer    '������� ��������
Public giRealCountZ As Integer    '�������� ������� ��������

Public gbDontStat   As Boolean    ' ���� �������� (������)


'��������� ��������
Public gbStopAGNKS  As Boolean    ' ���� �������� �����

'��������
Public gdTime       As Double    '����� ��������


Public giDVS        As Integer
Public gdPlot       As Double