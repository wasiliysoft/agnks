Attribute VB_Name = "AvgHelperModule"
Option Explicit

' ���������� ������ ��� ������� ������� �������� ��������
Private Const ticToCalcAvgRefuelingSpeed = 4
Private avgRefuelingSpeed As Double
Private last_gd������1 As Double

' ���������� ������ ��� ������� ������� �� ��������� ��������
Private Const ticToCalcLeftRefuelingTime = 20
Private pCarArr(ticToCalcLeftRefuelingTime) As Double
Private last_pCarSumm As Double
Private lastLeftRefuelingTime As Double

Function getAvgRefuelingSpeed() As Double
    If (nTimer1Counter Mod ticToCalcAvgRefuelingSpeed = 0) Then
      avgRefuelingSpeed = (gd������1 - last_gd������1) * (60000 / (frmStart.Timer1.Interval * ticToCalcAvgRefuelingSpeed))
      avgRefuelingSpeed = Round(avgRefuelingSpeed, 2)
      If avgRefuelingSpeed < 0 Then avgRefuelingSpeed = 0
      last_gd������1 = gd������1
    End If
    getAvgRefuelingSpeed = avgRefuelingSpeed
End Function


Function getLeftRefuelingTime() As Double
    Dim pCarSumm As Double
    if k5_isOpen Then
        pCarArr(nTimer1Counter Mod ticToCalcLeftRefuelingTime) = gnDif(4)
        pCarSumm = summArray(pCarArr) / ticToCalcLeftRefuelingTime  ' ���������
        if pCarSumm < gdUpLevel Then ' �������� �� ���������
          If (nTimer1Counter Mod ticToCalcLeftRefuelingTime = 0) Then
            If ((pCarSumm - last_pCarSumm) > 0) Then
              ' ���������� ����������� �� ���� �������� *
              ' �������� ���� ������ �� 1000 ��������� ������ � ������� * ���������� �����.
              lastLeftRefuelingTime = ((gdUpLevel - pCarSumm) / (pCarSumm - last_pCarSumm)) _
              * ((frmStart.Timer1.Interval / 1000) * ticToCalcLeftRefuelingTime)
              lastLeftRefuelingTime = Round(lastLeftRefuelingTime, 0)
            End If
            last_pCarSumm = pCarSumm
          End If
        Else  ' �������� ���������, ���� ���������
          lastLeftRefuelingTime = 0        
        End If
    Else ' ���� ������
        pCarArr(nTimer1Counter Mod ticToCalcLeftRefuelingTime) = 0
        lastLeftRefuelingTime = 0
    End If
    getLeftRefuelingTime = lastLeftRefuelingTime
End Function
