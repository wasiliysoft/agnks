Attribute VB_Name = "AvgHelperModule"
Option Explicit

Private avgRefuelingSpeed As Double
Private last_gdРасход1 As Double

Private Const ticToCalc = 4 ' количество тактов для расчета

Function getAvgRefuelingSpeed() As Double
    If (nTimer1Counter Mod ticToCalc = 0) Then
      avgRefuelingSpeed = (gdРасход1 - last_gdРасход1) * (60000 / (frmStart.Timer1.Interval * ticToCalc))
      avgRefuelingSpeed = Round(avgRefuelingSpeed, 2)
      last_gdРасход1 = gdРасход1
    End If
    getAvgRefuelingSpeed = avgRefuelingSpeed
End Function

