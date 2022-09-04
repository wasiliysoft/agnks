Attribute VB_Name = "AvgHelperModule"
Option Explicit

Private avgRefuelingSpeed As Double
Private last_gd������1 As Double

Private Const ticToCalc = 4 ' ���������� ������ ��� �������

Function getAvgRefuelingSpeed() As Double
    If (nTimer1Counter Mod ticToCalc = 0) Then
      avgRefuelingSpeed = (gd������1 - last_gd������1) * (60000 / (frmStart.Timer1.Interval * ticToCalc))
      avgRefuelingSpeed = Round(avgRefuelingSpeed, 2)
      last_gd������1 = gd������1
    End If
    getAvgRefuelingSpeed = avgRefuelingSpeed
End Function

