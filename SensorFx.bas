Attribute VB_Name = "SensorFx"
Public SMean(23) As Double

Public Function TrNumber(TNumber As Double, TDec As Long) As Double
TDec = Abs(TDec)
If TDec = 0 Then TDec = -1
If Not InStr(CStr(TNumber), "E") = 0 Then
TNumber = CDbl(Mid(TNumber, 1, InStr(CStr(TNumber), ",") + TDec + 2) / 10 ^ (CLng(Mid(CStr(TNumber), InStr(CStr(TNumber), "E") + 2, Len(TNumber)))))
End If
'---Add-----------------------
If InStr(CStr(TNumber), ",") = 0 Then GoTo sn
'-----------------------------
TrNumber = CDbl(Mid(TNumber, 1, InStr(CStr(TNumber), ",") + TDec))
Exit Function
sn:
TrNumber = CDbl(TNumber)
'Debug.Print CDbl(Mid(TNumber, 1, InStr(CStr(TNumber), ",") + TDec))
End Function

Public Function SensorRange(SIndex As Long, SInput As Long) As Double
Dim SVal As Double
Dim p As String
SVal = Round(SInput / 4095, 4)
'GoTo QL
Select Case SIndex
Case 0
SVal = TrNumber(-3.2371 + SVal ^ (1) * 15.8844 + SVal ^ (2) * -14.5134 + SVal ^ (3) * 25.4191 + SVal ^ (4) * -14.7596, 2)
Case 1
SVal = TrNumber(-32.499 + SVal ^ (1) * 196.682 + SVal ^ (2) * 30.16 + SVal ^ (3) * -43.679, 2)
Case 2
SVal = TrNumber(-31.358 + SVal ^ (1) * 166.352 + SVal ^ (2) * 100.303 + SVal ^ (3) * -84.968, 2)
Case 3
SVal = 0
Case 4
SVal = TrNumber(-3.0461 + SVal ^ (1) * 14.7716 + SVal ^ (2) * -10.6483 + SVal ^ (3) * 19.4589 + SVal ^ (4) * -11.6648, 2)
Case 5
SVal = TrNumber(-31.672 + SVal ^ (1) * 200.849 + SVal ^ (2) * 12.318 + SVal ^ (3) * -30.586, 2)
Case 6
SVal = TrNumber(-33.666 + SVal ^ (1) * 247.766 + SVal ^ (2) * -219.357 + SVal ^ (3) * 341.234 + SVal ^ (4) * -187.637, 2)
Case 7
SVal = TrNumber(-1.8756 + SVal ^ (1) * 13.4153 + SVal ^ (2) * -10.4675 + SVal ^ (3) * 17.3159 + SVal ^ (4) * -9.9519, 2)
Case 8
SVal = TrNumber(-25.46 + SVal ^ (1) * 158.149 + SVal ^ (2) * 16.159 + SVal ^ (3) * -28.69, 2)
Case 9
SVal = TrNumber(-33.122 + SVal ^ (1) * 245.487 + SVal ^ (2) * -213.03 + SVal ^ (3) * 340.199 + SVal ^ (4) * -190.352, 2)
Case 10
SVal = TrNumber(-32.367 + SVal ^ (1) * 196.475 + SVal ^ (2) * 24.232 + SVal ^ (3) * -39.104, 2)
Case 11
SVal = TrNumber(-34.284 + SVal ^ (1) * 233.521 + SVal ^ (2) * -114.378 + SVal ^ (3) * 161.332 + SVal ^ (4) * -97.182, 2)
Case 12
SVal = TrNumber(-34.291 + SVal ^ (1) * 243.422 + SVal ^ (2) * -197.28 + SVal ^ (3) * 332.984 + SVal ^ (4) * -196.55, 2)
Case 13
SVal = TrNumber(-33.848 + SVal ^ (1) * 238.255 + SVal ^ (2) * -164.909 + SVal ^ (3) * 283.101 + SVal ^ (4) * -172.781, 2)
Case 14
SVal = TrNumber(-34.255 + SVal ^ (1) * 238.302 + SVal ^ (2) * -164.902 + SVal ^ (3) * 275.373 + SVal ^ (4) * -165.302, 2)
Case 15
SVal = TrNumber(-3.3174 + SVal ^ (1) * 16.2226 + SVal ^ (2) * -12.2655 + SVal ^ (3) * 20.5197 + SVal ^ (4) * -12.2447, 2)
Case 16
SVal = TrNumber(-3.1519 + SVal ^ (1) * 14.8175 + SVal ^ (2) * -7.1292 + SVal ^ (3) * 11.3764 + SVal ^ (4) * -6.9538, 2)
'SVal = Round(-3.1519 + SVal ^ (1) * 14.8175 + SVal ^ (2) * -7.1292 + SVal ^ (3) * 11.3764 + SVal ^ (4) * -6.9538, 1)
Case 17
SVal = TrNumber(-3.0096 + SVal ^ (1) * 14.4714 + SVal ^ (2) * -8.6875 + SVal ^ (3) * 16.1977 + SVal ^ (4) * -10.0252, 2)
'SVal = Round(-3.0496 + SVal ^ (1) * 14.4714 + SVal ^ (2) * -8.6875 + SVal ^ (3) * 16.1977 + SVal ^ (4) * -10.0252, 1)
Case 18
SVal = TrNumber(-33.406 + SVal ^ (1) * 242.321 + SVal ^ (2) * -188.413 + SVal ^ (3) * 303.612 + SVal ^ (4) * -174.565, 2)
Case 19
SVal = TrNumber(-91.901 + SVal ^ (1) * -1216.7 + SVal ^ (2) * 20319 + SVal ^ (3) * -22493 + SVal ^ (4) * 8040.9, 0)
Case 20
SVal = 0
Case 21
SVal = TrNumber(-33.553 + SVal ^ (1) * 245.518 + SVal ^ (2) * -203.31 + SVal ^ (3) * 326.697 + SVal ^ (4) * -186.111, 2)
Case 22
SVal = TrNumber(-34.291 + SVal ^ (1) * 242.552 + SVal ^ (2) * -180.651 + SVal ^ (3) * 301.889 + SVal ^ (4) * -180.253, 2)
Case 23
SVal = TrNumber(-994.16 + SVal ^ (1) * 6834.25 + SVal ^ (2) * -3909.16 + SVal ^ (3) * 6804.6 + SVal ^ (4) * -4257.19, 0)
End Select
QH:
'SensorRange = SVal
'-1 or 2 decimals-----------------
If SIndex = 19 Or SIndex = 23 Then
SensorRange = TrNumber((SVal + SMean(SIndex)) / 2, 0)
ElseIf SIndex = 4 Or SIndex = 16 Or SIndex = 17 Then
SensorRange = TrNumber((SVal + SMean(SIndex)) / 2, 2)
Else
SensorRange = TrNumber((SVal + SMean(SIndex)) / 2, 1)
End If
SMean(SIndex) = SVal
QL:
'SensorRange = SVal
End Function
