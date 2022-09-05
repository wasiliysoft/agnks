Attribute VB_Name = "MetanCounter_Helper"
Option Explicit

'�������� ������� �������� ������� ���� (��������)
Private Declare Sub ResetExpenseCounter Lib "MetanCounter" Alias "#1" (ByVal i As Long)
Declare Sub AddSensorsData Lib "MetanCounter" Alias "#2" (ByVal i As Long, ByVal _
        p1 As Double, ByVal t1 As Double, ByVal p2 As Double, ByVal d As Double, ByVal _
        coef As Double, ByVal CorrExp As Double)
Private Declare Function GetMassExpense Lib "MetanCounter" Alias "#4" (ByVal i As Long) As Double
Private Declare Function GetMass Lib "MetanCounter" Alias "#5" (ByVal i As Long) As Double
Private Declare Function GetTimeCounter Lib "MetanCounter" Alias "#6" (ByVal i As Long) As Double
Declare Sub StartOutput Lib "MetanCounter" Alias "#7" (ByVal i As Long)
Declare Sub StopOutput Lib "MetanCounter" Alias "#8" (ByVal i As Long)

Private gTime2 As Double
Private gMass2 As Double

Sub ResetExpenseCounter_2()
    If isDebug Then
        gTime2 = 0
        gMass2 = 0
    Else
        ResetExpenseCounter (2)
    End If
End Sub

Function GetTimeCounter_2() As Double
    If isDebug Then
        gTime2 = gTime2 + 0.5
        GetTimeCounter_2 = gTime2
    Else
        GetTimeCounter_2 = GetTimeCounter(2)
    End If
End Function

Function GetMass_2() As Double
    If isDebug Then
        If giStage2 = 9 Then
            gMass2 = gMass2 + 0.15
        ElseIf giStage2 = 4 Then
            gMass2 = gMass2 + 0.2
        Else
            gMass2 = 0
        End If
        GetMass_2 = gMass2
    Else
        GetMass_2 = GetMass(2)
    End If
End Function

Function GetMassExpense_2() As Double
    GetMassExpense_2 = GetMassExpense(2)
End Function
