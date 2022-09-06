Attribute VB_Name = "AvgHelperModule"
Option Explicit

' количество тактов для расчета времени до окончания заправки
Private Const ticToCalcLeftRefuelingTime = 20
Private pCarArr(ticToCalcLeftRefuelingTime) As Double
Private last_pCarSumm As Double
Private lastLeftRefuelingTime As Double


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
