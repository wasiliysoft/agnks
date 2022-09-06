Attribute VB_Name = "AvgHelperModule"
Option Explicit

' ���������� ������ ��� ������� ������� �� ��������� ��������
Private Const ticToCalcLeftRefuelingTime = 20
Private pCarArr(ticToCalcLeftRefuelingTime) As Double
Private last_pCarSumm As Double
Private lastLeftRefuelingTime As Double


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
