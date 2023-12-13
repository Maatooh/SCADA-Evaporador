Attribute VB_Name = "SaveFx"
Public Function SaveSwitch()
Dim SBuff As String
'--Switch0--------------------------------
SBuff = SBuff & "[Switch0]:"
For S = Server.Switch0.LBound To Server.Switch0.UBound
Select Case Server.Switch0(S).Caption
Case "ON"
SBuff = SBuff & 1 & ";"
Case "OFF"
SBuff = SBuff & 0 & ";"
End Select
Next S
SBuff = Mid(SBuff, 1, Len(SBuff) - 1) & vbCrLf
'-----------------------------------------
'--Switch1--------------------------------
SBuff = SBuff & "[Switch1]:"
For S = Server.Switch1.LBound To Server.Switch1.UBound
Select Case Server.Switch1(S).Caption
Case "ON"
SBuff = SBuff & 1 & ";"
Case "OFF"
SBuff = SBuff & 0 & ";"
End Select
Next S
SBuff = Mid(SBuff, 1, Len(SBuff) - 1) & vbCrLf
'-----------------------------------------
'--Switch2--------------------------------
SBuff = SBuff & "[Switch2]:"
For S = Server.Switch2.LBound To Server.Switch2.UBound
Select Case Server.Switch2(S).Caption
Case "ON"
SBuff = SBuff & 1 & ";"
Case "OFF"
SBuff = SBuff & 0 & ";"
End Select
Next S
SBuff = Mid(SBuff, 1, Len(SBuff) - 1) & vbCrLf
'-----------------------------------------
'--Switch3--------------------------------
SBuff = SBuff & "[Switch3]:"
For S = Server.Switch3.LBound To Server.Switch3.UBound
Select Case Server.Switch3(S).Caption
Case "ON"
SBuff = SBuff & 1 & ";"
Case "OFF"
SBuff = SBuff & 0 & ";"
End Select
Next S
SBuff = Mid(SBuff, 1, Len(SBuff) - 1) & vbCrLf
'-----------------------------------------
'--Others--------------------------------
SBuff = SBuff & "[Others]:"
'---Alarm
Select Case Server.BAlarm.Caption
Case "ON"
SBuff = SBuff & 1 & ";"
Case "OFF"
SBuff = SBuff & 0 & ";"
End Select
'---Ev
Select Case Server.EvAgua.Caption
Case "ON"
SBuff = SBuff & 1 & ";"
Case "OFF"
SBuff = SBuff & 0 & ";"
End Select
'---EvManual/Auto
Select Case Server.EvAM.Caption
Case "Auto"
SBuff = SBuff & 1 & ";"
Case "Man"
SBuff = SBuff & 0 & ";"
End Select
'---
SBuff = Mid(SBuff, 1, Len(SBuff) - 1) & vbCrLf
'-----------------------------------------
'--PID-SetPoint---------------------------
SBuff = SBuff & "[PID]:"
For S = LBound(ControlFx.SensorParams) To UBound(ControlFx.SensorParams)
SBuff = SBuff & ControlFx.SensorParams(S) & ";"
Next S
SBuff = Mid(SBuff, 1, Len(SBuff) - 1) & vbCrLf
'-----------------------------------------
Open App.Path & "\SaveSwitches.dat" For Output As #1
Print #1, SBuff
Close #1
'Debug.Print SBuff
End Function

Public Function LoadSwitch()
On Error GoTo QH
Dim LBuff As String
If Dir(App.Path & "\SaveSwitches.dat", vbArchive) = vbNullString Then GoTo QH
Open App.Path & "\SaveSwitches.dat" For Input As #1
LBuff = Input$(LOF(1), #1)
Close #1
'------------------------------
Dim Buffers() As String
Dim Buff0() As String
Dim Buff1() As String
Dim Buff2() As String
Dim Buff3() As String
Dim OBuff() As String
Dim SBuff() As String
'------------------------------
Buffers = Split(LBuff, vbCrLf)
Buff0 = Split(Replace(Buffers(0), "[Switch0]:", vbNullString), ";")
Buff1 = Split(Replace(Buffers(1), "[Switch1]:", vbNullString), ";")
Buff2 = Split(Replace(Buffers(2), "[Switch2]:", vbNullString), ";")
Buff3 = Split(Replace(Buffers(3), "[Switch3]:", vbNullString), ";")
OBuff = Split(Replace(Buffers(4), "[Others]:", vbNullString), ";")
SBuff = Split(Replace(Buffers(5), "[PID]:", vbNullString), ";")
'-----------------------------
'-----Load in Switch 0
For f = Server.Switch0.LBound To Server.Switch0.UBound
Select Case Buff0(f)
Case 1
Server.Switch0(f).Caption = "ON"
Server.Switch0(f).BackColor = &H80FF80
Case 0
Server.Switch0(f).Caption = "OFF"
Server.Switch0(f).BackColor = &H8080FF
End Select
Next f
'-----Load in Switch 1
For f = Server.Switch1.LBound To Server.Switch1.UBound
Select Case Buff1(f)
Case 1
Server.Switch1(f).Caption = "ON"
Server.Switch1(f).BackColor = &H80FF80
Case 0
Server.Switch1(f).Caption = "OFF"
Server.Switch1(f).BackColor = &H8080FF
End Select
Next f
'-----Load in Switch 2
For f = Server.Switch2.LBound To Server.Switch2.UBound
Select Case Buff2(f)
Case 1
Server.Switch2(f).Caption = "ON"
Server.Switch2(f).BackColor = &H80FF80
Case 0
Server.Switch2(f).Caption = "OFF"
Server.Switch2(f).BackColor = &H8080FF
End Select
Next f
'-----Load in Switch 3
For f = Server.Switch3.LBound To Server.Switch3.UBound
Select Case Buff3(f)
Case 1
Server.Switch3(f).Caption = "ON"
Server.Switch3(f).BackColor = &H80FF80
Case 0
Server.Switch3(f).Caption = "OFF"
Server.Switch3(f).BackColor = &H8080FF
End Select
Next f
'----Load in Others
Select Case OBuff(0)
Case 1
Server.BAlarm.Caption = "ON"
Server.BAlarm.BackColor = &H80FF80
Case 0
Server.BAlarm.Caption = "OFF"
Server.BAlarm.BackColor = &H8080FF
End Select
Select Case OBuff(1)
Case 1
Server.EvAgua.Caption = "ON"
Server.EvAgua.BackColor = &H80FF80
Case 0
Server.EvAgua.Caption = "OFF"
Server.EvAgua.BackColor = &H8080FF
End Select
Select Case OBuff(2)
Case 1
Server.EvAM.Caption = "Auto"
Server.EvAM.BackColor = &H80FF80
Case 0
Server.EvAM.Caption = "Man"
Server.EvAM.BackColor = &H8080FF
End Select
'----Load in PID
For f = LBound(ControlFx.SensorParams) To UBound(ControlFx.SensorParams)
ControlFx.SensorParams(f) = SBuff(f)
Next f
QH:
End Function
