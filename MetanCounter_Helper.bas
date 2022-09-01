Attribute VB_Name = "MetanCounter_Helper"
Option Explicit

'�������� ������� �������� ������� ���� (��������)
Declare Sub ResetExpenseCounter Lib "MetanCounter" Alias "#1" (ByVal i As Long)
Declare Sub AddSensorsData Lib "MetanCounter" Alias "#2" (ByVal i As Long, ByVal _
        p1 As Double, ByVal t1 As Double, ByVal p2 As Double, ByVal d As Double, ByVal _
        coef As Double, ByVal CorrExp As Double)
Private  Declare Function GetMassExpense Lib "MetanCounter" Alias "#4" (ByVal i As Long) As Double
Private  Declare Function GetMass Lib "MetanCounter" Alias "#5" (ByVal i As Long) As Double
Private  Declare Function GetTimeCounter Lib "MetanCounter" Alias "#6" (ByVal i As Long) As Double
Declare Sub StartOutput Lib "MetanCounter" Alias "#7" (ByVal i As Long)
Declare Sub StopOutput Lib "MetanCounter" Alias "#8" (ByVal i As Long)


Function GetTimeCounter_2() As Double
    GetTimeCounter_2 = GetTimeCounter(2)
End Function

Function GetMass_1() As Double
    GetMass_1 = GetMass(1)
End Function

Function GetMass_2() As Double
    GetMass_2 = GetMass(2)
End Function

Function GetMassExpense_2() As Double
    GetMassExpense_2 = GetMassExpense(2)
End Function