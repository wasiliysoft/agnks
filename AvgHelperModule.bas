Attribute VB_Name = "AvgHelperModule"
Option Explicit

' количество тактов для расчета средней скорости заправки
Private Const ticToCalcAvgRefuelingSpeed = 4
Private avgRefuelingSpeed As Double
Private last_gdРасход1 As Double

' количество тактов для расчета времени до окончания заправки
Private Const ticToCalcLeftRefuelingTime = 20
Private pCarArr(ticToCalcLeftRefuelingTime) As Double
Private last_pCarSumm As Double
Private lastLeftRefuelingTime As Double

Function getAvgRefuelingSpeed() As Double
    If (nTimer1Counter Mod ticToCalcAvgRefuelingSpeed = 0) Then
      avgRefuelingSpeed = (gdРасход1 - last_gdРасход1) * (60000 / (frmStart.Timer1.Interval * ticToCalcAvgRefuelingSpeed))
      avgRefuelingSpeed = Round(avgRefuelingSpeed, 2)
      If avgRefuelingSpeed < 0 Then avgRefuelingSpeed = 0
      last_gdРасход1 = gdРасход1
    End If
    getAvgRefuelingSpeed = avgRefuelingSpeed
End Function


Function getLeftRefuelingTime() As Double
    Dim pCarSumm As Double
    if k5_isOpen Then
        pCarArr(nTimer1Counter Mod ticToCalcLeftRefuelingTime) = gnDif(4)
        pCarSumm = summArray(pCarArr) / ticToCalcLeftRefuelingTime  ' Усредняем
        if pCarSumm < gdUpLevel Then ' Давление не превышено
          If (nTimer1Counter Mod ticToCalcLeftRefuelingTime = 0) Then
            If ((pCarSumm - last_pCarSumm) > 0) Then
              ' Количество промежутков до МАКС давления *
              ' Интервал тика делить на 1000 потомучто таймер в милисек * количество тиков.
              lastLeftRefuelingTime = ((gdUpLevel - pCarSumm) / (pCarSumm - last_pCarSumm)) _
              * ((frmStart.Timer1.Interval / 1000) * ticToCalcLeftRefuelingTime)
              lastLeftRefuelingTime = Round(lastLeftRefuelingTime, 0)
            End If
            last_pCarSumm = pCarSumm
          End If
        Else  ' Давление превышено, авто заправлен
          lastLeftRefuelingTime = 0        
        End If
    Else ' Кран закрыт
        pCarArr(nTimer1Counter Mod ticToCalcLeftRefuelingTime) = 0
        lastLeftRefuelingTime = 0
    End If
    getLeftRefuelingTime = lastLeftRefuelingTime
End Function
