Attribute VB_Name = "Module1"
'��� ���������� ���������� ��������� �������
Option Explicit

'����������� ��������
Global Const ggMinPress = 197    ' ����������� �������� � �������������
Global Const A0 = 0
Global Const A1 = 256
Global Const B0 = 1          '�����
Global Const B1 = 257
Global Const C0 = 2
Global Const C1 = 258
Global Const conHwndTopmost = -1
Global Const conSwpNoActivate = &H10
Global Const conSwpShowWindow = &H40


'�������� ������� ��� ����� ACL-8113
'*********************************************************************************
'      The Declare of ISO813.DLL for ISO813 AD Card
'*********************************************************************************

'****** define the error number *******/
Global Const ISO813_NoError = 0
Global Const ISO813_CheckBoardError = 1
Global Const ISO813_DriverOpenError = 2
Global Const ISO813_DriverNoOpen = 3
Global Const ISO813_AdError = 4
Global Const ISO813_OtherError = 5
Global Const ISO813_GetDriverVersionError = 6
Global Const ISO813_TimeOutError = &HFFFF

' Function of Driver
Declare Function ISO813_DriverInit Lib "ISO813.DLL" () As Integer
Declare Sub ISO813_DriverClose Lib "ISO813.DLL" ()

' Function of AD
Declare Function ISO813_AD_Float Lib "ISO813.DLL" (ByVal wBase As Integer, ByVal wChannel As Integer, _
        ByVal wGainCode As Integer, ByVal wBipolar As Integer, _
        ByVal wJmp10v As Integer) As Single




'�������� ������� �������� ������� ���� (��������)
Declare Sub ResetExpenseCounter Lib "MetanCounter" Alias "#1" (ByVal i As Long)


Declare Sub AddSensorsData Lib "MetanCounter" Alias "#2" (ByVal i As Long, ByVal _
        p1 As Double, ByVal t1 As Double, ByVal p2 As Double, ByVal d As Double, ByVal _
        coef As Double, ByVal CorrExp As Double)
Declare Function GetMassExpense Lib "MetanCounter" Alias "#4" (ByVal i As Long) As Double
Declare Function GetMass Lib "MetanCounter" Alias "#5" (ByVal i As Long) As Double
Declare Function GetTimeCounter Lib "MetanCounter" Alias "#6" (ByVal i As Long) As Double
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Declare Sub StartOutput Lib "MetanCounter" Alias "#7" (ByVal i As Long)
Declare Sub StopOutput Lib "MetanCounter" Alias "#8" (ByVal i As Long)



'���������� ����������������� ������
'������ ������� + ��������
Public Type Sensor
    Data            As Integer
    Note            As String
End Type



'���������� ���������� � ��������
Public gKv          As Double
Public gKi          As Double
Public gKt          As Double
Public gKp          As Double    '��� ��1.1 � ��1.2
Public gKp_1        As Double    '��� ���������
Public gKi_1        As Double
Public gKn          As Double


'��� ������� ����
Public gd��1        As Double
Public gd��2        As Double


'��� ��������
Public gbOnlyAkk    As Boolean    '���� �������� ������ �� �������������
'������ ����� �� �������� ���
Public giStage      As Integer    '����� ��������
Public giStage1     As Integer    '�������� ������������� �����
Public giStage2     As Integer    '�������� ����� ��������
Public gbFrmShow    As Boolean
Public gbCmdStart   As Boolean    '��� ������ ����\��������
Public gbAkkum      As Boolean
Public giTrigger    As Integer    ' ��� ����� �������
Public gsMsg        As String
Public Car          As Integer
Public gb��������   As Boolean    '���������� ����� ����� � ������� ����� ������� ������
'���� �� �������� ������ ����������
Public gd������1    As Double    '������ �� �������� ����������
Public giMain������ As Integer    '���� ��� �������� ������� ����� ����:
' 1 - ���������
' 0 - ������
'-1 - ��������

Public gdK          As Double    '����������� �����������
Public gdRashAkkEnd As Double    '������ ������ �� �������� ���������� ����� �� �������������

' ���������� ��� ����������
Public glAver       As Long    ' ������ ����� ��� ����������
Public glCounter    As Long
Public sum(31)      As Double


Public gl���������  As Integer
Public gn48DIO(5)   As Long    '��������� ��������� ����� PET-48DIO
Public gn������(48) As Sensor    '��������� �������� �� ������ TB-24P � TB-16P8R
Public ggACL8113(31) As Double   '��������� �������� ����� 8113
Public gnDif(31)    As Double    ' ��� ������������� ��������(� ���� � ���� ������)
Public gs���������  As String
Public gl��������   As Long
Public gl�����      As Long
Public gla�����     As Long









Public CN           As Integer


Public gbRunDVS     As Boolean

Public gdUpLevel    As Double
Public giChanel     As Integer

Public Type MyRecType
    dt              As Date
    IR1             As Double
    IR2             As Double
    Motor           As Long
End Type

Public FileHandle   As Integer    '������������� ����� � ����������
Public MotorCount   As Long    '������� ������������
Public GMC          As Long    '���������� �������

Public StatDB As Database, StatWS As Workspace
Public StatRS       As Recordset
Public SelectRS     As Recordset

'������ �� ����� ���� (� ������� �� ����)
'0-�� ������� - �������
Public gdaStat1(90) As MyRecType    '������ ������ �� ��������� �� ����
Public gdaStat2(31) As MyRecType    '������ ������ �� ��������� �� �����
Public gdaStat3(12) As MyRecType    '������ ������ �� ��������� �� ���
Public gdaStat4(100) As MyRecType    '������ ������ �� �����

Public gDateRec     As Date    '���� ��������� ������
Public giCountZ     As Integer    '������� ��������
Public giRealCountZ As Integer    '�������� ������� ��������

Public gbDontStat   As Boolean    ' ���� �������� (������)
Public gbHandControl As Boolean    ' ���� ������� ����������


'��������� ��������
Public gbStopAGNKS  As Boolean    ' ���� �������� �����
Public gbFireDVS    As Boolean    '����� � ������ ���
Public gbFireTech   As Boolean    '����� � ���. ������


'��������
Public gdTime       As Double    '����� ��������

'��������� ��� secret file
Public Type pswd
    PC              As Double
    pwd             As String * 7
End Type
Public Password     As String
Public giDVS        As Integer
Public gdPlot       As Double