Attribute VB_Name = "ControlFx"
Public SensorParams(23) As Long

Public Sub ClassParamsSensor(CMD As String)
On Error GoTo QH
Dim SIndex As Long
Dim SValue As String
Dim SValues() As String
SIndex = CLng(Mid(CMD, InStrRev(CMD, "S") + 1, InStr(CMD, "]") - InStr(CMD, "S") - 1))
SValue = CStr(Mid(CMD, InStrRev(CMD, "]:") + 2, Len(CMD)))
SValue = Mid(SValue, InStrRev(SValue, "[Ri]") + 4, Len(SValue))
'Debug.Print CMD
SValues = Split(SValue, "|")
'---------Sensores----------
Select Case SIndex
Case 0
For S = 0 To 11
Server.Sensor(S).Caption = Round(SensorFx.SensorRange(CLng(S), CLng(SValues(S))), 4)
Next S
Case 1
For S = 0 To 11
Server.Sensor(S + 12).Caption = Round(SensorFx.SensorRange(CLng(S + 12), CLng(SValues(S))), 4)
Next S
End Select
'---------EV Agua----------
If Server.EvAM.Caption = "Auto" And SIndex = 0 Then
Select Case SValues(UBound(SValues))
'--------------
Case 0
If Server.EvAgua.Caption = "OFF" Then
Server.EvAgua.Caption = "ON"
Server.EvAgua.BackColor = &H80FF80
Call ControlFx.SendValvStatus("[CMD V2]:" & ControlFx.SwitchState(2), "vDev2")
End If
Case 1
If Server.EvAgua.Caption = "ON" Then
Server.EvAgua.Caption = "OFF"
Server.EvAgua.BackColor = &H8080FF
Call ControlFx.SendValvStatus("[CMD V2]:" & ControlFx.SwitchState(2), "vDev2")
End If
'---------------
End Select
End If
'Server.Sensor(SIndex) = SValue
QH:
End Sub

Public Sub SendValvStatus(CMD As String, DevNameID As String)
For w = 0 To UBound(DevicesFx.DevName)
If Server.WServer(w).State = sckConnected And DevNameID = DevicesFx.DevName(w) Then
Server.WServer(w).SendData CMD
End If
Next w
End Sub

Public Function SwitchState(SwitchSets As Long) As String
Dim SwitchBuff As String
Select Case SwitchSets
Case 0
'---Switch 0---
For j = 0 To Server.Switch0.UBound
If Server.Switch0(j).Caption = "OFF" Then SwitchBuff = SwitchBuff & 0
If Server.Switch0(j).Caption = "ON" Then SwitchBuff = SwitchBuff & 1
Next j
'--------------
Case 1
'---Switch 1---
For j = 0 To Server.Switch1.UBound
If Server.Switch1(j).Caption = "OFF" Then SwitchBuff = SwitchBuff & 0
If Server.Switch1(j).Caption = "ON" Then SwitchBuff = SwitchBuff & 1
Next j
'--------------
Case 2
'---Switch 2---
For j = 0 To Server.Switch2.UBound
If Server.Switch2(j).Caption = "OFF" Then SwitchBuff = SwitchBuff & 0
If Server.Switch2(j).Caption = "ON" Then SwitchBuff = SwitchBuff & 1
Next j
'-Electro Valvula
If Server.EvAgua.Caption = "OFF" Then SwitchBuff = SwitchBuff & 0
If Server.EvAgua.Caption = "ON" Then SwitchBuff = SwitchBuff & 1
'-Alarma Baliza
If Server.BAlarm.Caption = "OFF" Then SwitchBuff = SwitchBuff & 0
If Server.BAlarm.Caption = "ON" Then SwitchBuff = SwitchBuff & 1
'--------------
Case 3
'---Switch 3---
For j = 0 To Server.Switch3.UBound
If Server.Switch3(j).Caption = "OFF" Then SwitchBuff = SwitchBuff & 0
If Server.Switch3(j).Caption = "ON" Then SwitchBuff = SwitchBuff & 1
Next j
'--------------
End Select
SwitchState = SwitchBuff
End Function

Public Sub RefreshState()
Call ControlFx.SendValvStatus("[CMD V0]:" & ControlFx.SwitchState(0), "vDev0")
Call DevicesFx.delay(0.5)
Call ControlFx.SendValvStatus("[CMD V1]:" & ControlFx.SwitchState(1), "vDev1")
Call DevicesFx.delay(0.5)
Call ControlFx.SendValvStatus("[CMD V2]:" & ControlFx.SwitchState(2), "vDev2")
Call DevicesFx.delay(0.5)
Call ControlFx.SendValvStatus("[CMD V3]:" & ControlFx.SwitchState(3), "vDev3")
Call DevicesFx.delay(0.5)
Call ControlFx.SetPID("P0", "PDev0")
End Sub

Public Sub SetLeds(CMD As String)
'Debug.Print CMD
Dim SIndex As Long
Dim SValue As String
Dim SValues() As String
SIndex = CLng(Mid(CMD, InStrRev(CMD, "L") + 1, InStr(CMD, "]") - InStr(CMD, "L") - 1))
SValue = CStr(Mid(CMD, InStrRev(CMD, "]:") + 2, Len(CMD)))
ReDim Preserve SValues(18)
For l = 0 To 18
SValues(l) = Mid(SValue, (l + 1), 1)
Next l
'---------------
Select Case SIndex
Case 0
'-----------------
For m = 0 To 18
If SValues(m) = 0 Then
Server.LEDR(m).BackColor = &HDBDBFE
End If
If SValues(m) = 1 Then
Server.LEDR(m).BackColor = &HFF&
End If
Next m
'-----------------
Case 1
'-----------------
For m = 0 To 18
If SValues(m) = 0 Then
Server.LEDR(m + 19).BackColor = &HDBDBFE
End If
If SValues(m) = 1 Then
Server.LEDR(m + 19).BackColor = &HFF&
End If
Next m
'-----------------
End Select
End Sub

Public Function ReadPID() As String
Dim SBuff As String
For p = 0 To Server.Sensor.UBound
If Server.Sensor(p).BackColor = &H80FF& Then
SBuff = SBuff & (SensorParams(p) / 100) * 255 & "|"
End If
Next p
SBuff = Mid(SBuff, 1, Len(SBuff) - 1)
ReadPID = SBuff
End Function

Public Sub SetPID(CMDName As String, DevNameID As String)
For w = 0 To UBound(DevicesFx.DevName)
If Server.WServer(w).State = sckConnected And DevNameID = DevicesFx.DevName(w) Then
Server.WServer(w).SendData "[CMD " & CMDName & "]:" & ReadPID
'Debug.Print ReadPID
End If
Next w
End Sub
