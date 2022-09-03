Attribute VB_Name = "Utils"
Option Explicit

Public Function formatSecToHHMMSS(ByVal s) As String
   Dim d As Date
   d = DateAdd("s", s, d)
   formatSecToHHMMSS = Format(d, "hh:nn:ss")
End Function

Public Function getP_As_Percent(ByVal currentP) As Integer
    Dim i as Integer
    i = 100 * (currentP / gdUpLevel)
    If (i > 100) Then
        i = 100
    ElseIf (i < 0) Then
        i = 0
    End If
    getP_As_Percent = i
End Function