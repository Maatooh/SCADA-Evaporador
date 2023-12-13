Attribute VB_Name = "LedFx"
'&HF4FFE5 OFF &HFF00& ON

Public Sub LedFunctionsS0(LIndex As Long)
Select Case LIndex
Case 0
Case 1
Case 2
Case 3
If Server.Switch0(LIndex).Caption = "ON" Then
Server.LEDG(13).BackColor = &HFF00&
ElseIf Server.Switch0(LIndex).Caption = "OFF" Then
Server.LEDG(13).BackColor = &HF4FFE5
End If
Case 4
If Server.Switch0(LIndex).Caption = "ON" Then
Server.LEDG(14).BackColor = &HFF00&
ElseIf Server.Switch0(LIndex).Caption = "OFF" Then
Server.LEDG(14).BackColor = &HF4FFE5
End If
Case 5
If Server.Switch0(LIndex).Caption = "ON" Then
Server.LEDG(0).BackColor = &HFF00&
ElseIf Server.Switch0(LIndex).Caption = "OFF" Then
Server.LEDG(0).BackColor = &HF4FFE5
End If
Case 6
If Server.Switch0(LIndex).Caption = "ON" Then
Server.LEDG(1).BackColor = &HFF00&
ElseIf Server.Switch0(LIndex).Caption = "OFF" Then
Server.LEDG(1).BackColor = &HF4FFE5
End If
Case 7
If Server.Switch0(LIndex).Caption = "ON" Then
Server.LEDG(11).BackColor = &HFF00&
ElseIf Server.Switch0(LIndex).Caption = "OFF" Then
Server.LEDG(11).BackColor = &HF4FFE5
End If
Case 8
If Server.Switch0(LIndex).Caption = "ON" Then
Server.LEDG(12).BackColor = &HFF00&
ElseIf Server.Switch0(LIndex).Caption = "OFF" Then
Server.LEDG(12).BackColor = &HF4FFE5
End If
Case 9
If Server.Switch0(LIndex).Caption = "ON" Then
Server.LEDG(31).BackColor = &HFF00&
ElseIf Server.Switch0(LIndex).Caption = "OFF" Then
Server.LEDG(31).BackColor = &HF4FFE5
End If
Case 10
If Server.Switch0(LIndex).Caption = "ON" Then
Server.LEDG(32).BackColor = &HF4FFE5
Server.LEDG(33).BackColor = &HFF00&
ElseIf Server.Switch0(LIndex).Caption = "OFF" Then
Server.LEDG(32).BackColor = &HFF00&
Server.LEDG(33).BackColor = &HF4FFE5
End If
Case 11
If Server.Switch0(LIndex).Caption = "ON" Then
Server.LEDG(53).BackColor = &HFF00&
ElseIf Server.Switch0(LIndex).Caption = "OFF" Then
Server.LEDG(53).BackColor = &HF4FFE5
End If
Case 12
If Server.Switch0(LIndex).Caption = "ON" Then
Server.LEDG(59).BackColor = &HFF00&
ElseIf Server.Switch0(LIndex).Caption = "OFF" Then
Server.LEDG(59).BackColor = &HF4FFE5
End If
Case 13
If Server.Switch0(LIndex).Caption = "ON" Then
Server.LEDG(56).BackColor = &HFF00&
ElseIf Server.Switch0(LIndex).Caption = "OFF" Then
Server.LEDG(56).BackColor = &HF4FFE5
End If
Case 14
If Server.Switch0(LIndex).Caption = "ON" Then
Server.LEDG(58).BackColor = &HFF00&
ElseIf Server.Switch0(LIndex).Caption = "OFF" Then
Server.LEDG(58).BackColor = &HF4FFE5
End If
Case 15
If Server.Switch0(LIndex).Caption = "ON" Then
Server.LEDG(57).BackColor = &HFF00&
ElseIf Server.Switch0(LIndex).Caption = "OFF" Then
Server.LEDG(57).BackColor = &HF4FFE5
End If
Case 16
If Server.Switch0(LIndex).Caption = "ON" Then
Server.LEDG(74).BackColor = &HF4FFE5
Server.LEDG(73).BackColor = &HFF00&
ElseIf Server.Switch0(LIndex).Caption = "OFF" Then
Server.LEDG(74).BackColor = &HFF00&
Server.LEDG(73).BackColor = &HF4FFE5
End If
Case 17
If Server.Switch0(LIndex).Caption = "ON" Then
Server.LEDG(72).BackColor = &HF4FFE5
Server.LEDG(71).BackColor = &HFF00&
ElseIf Server.Switch0(LIndex).Caption = "OFF" Then
Server.LEDG(72).BackColor = &HFF00&
Server.LEDG(71).BackColor = &HF4FFE5
End If
Case 18
If Server.Switch0(LIndex).Caption = "ON" Then
Server.LEDG(70).BackColor = &HF4FFE5
Server.LEDG(69).BackColor = &HFF00&
ElseIf Server.Switch0(LIndex).Caption = "OFF" Then
Server.LEDG(70).BackColor = &HFF00&
Server.LEDG(69).BackColor = &HF4FFE5
End If
Case 19
If Server.Switch0(LIndex).Caption = "ON" Then
Server.LEDG(67).BackColor = &HF4FFE5
Server.LEDG(68).BackColor = &HFF00&
Server.LEDG(71).BackColor = &HFF00&
ElseIf Server.Switch0(LIndex).Caption = "OFF" Then
Server.LEDG(67).BackColor = &HFF00&
Server.LEDG(68).BackColor = &HF4FFE5
End If
End Select
End Sub
Public Sub LedFunctionsS1(LIndex As Long)
Select Case LIndex
Case 0
If Server.Switch1(LIndex).Caption = "ON" Then
Server.LEDG(2).BackColor = &HFF00&
ElseIf Server.Switch1(LIndex).Caption = "OFF" Then
Server.LEDG(2).BackColor = &HF4FFE5
End If
Case 1
If Server.Switch1(LIndex).Caption = "ON" Then
Server.LEDG(5).BackColor = &HFF00&
ElseIf Server.Switch1(LIndex).Caption = "OFF" Then
Server.LEDG(5).BackColor = &HF4FFE5
End If
Case 2
If Server.Switch1(LIndex).Caption = "ON" Then
Server.LEDG(38).BackColor = &HFF00&
ElseIf Server.Switch1(LIndex).Caption = "OFF" Then
Server.LEDG(38).BackColor = &HF4FFE5
End If
Case 3
If Server.Switch1(LIndex).Caption = "ON" Then
Server.LEDG(7).BackColor = &HFF00&
ElseIf Server.Switch1(LIndex).Caption = "OFF" Then
Server.LEDG(7).BackColor = &HF4FFE5
End If
Case 4
If Server.Switch1(LIndex).Caption = "ON" Then
Server.LEDG(75).BackColor = &HFF00&
ElseIf Server.Switch1(LIndex).Caption = "OFF" Then
Server.LEDG(75).BackColor = &HF4FFE5
End If
Case 5
If Server.Switch1(LIndex).Caption = "ON" Then
Server.LEDG(62).BackColor = &HFF00&
ElseIf Server.Switch1(LIndex).Caption = "OFF" Then
Server.LEDG(62).BackColor = &HF4FFE5
End If
Case 6
If Server.Switch1(LIndex).Caption = "ON" Then
Server.LEDG(60).BackColor = &HFF00&
ElseIf Server.Switch1(LIndex).Caption = "OFF" Then
Server.LEDG(60).BackColor = &HF4FFE5
End If
Case 7
If Server.Switch1(LIndex).Caption = "ON" Then
Server.LEDG(63).BackColor = &HFF00&
ElseIf Server.Switch1(LIndex).Caption = "OFF" Then
Server.LEDG(63).BackColor = &HF4FFE5
End If
Case 8
If Server.Switch1(LIndex).Caption = "ON" Then
Server.LEDG(6).BackColor = &HFF00&
ElseIf Server.Switch1(LIndex).Caption = "OFF" Then
Server.LEDG(6).BackColor = &HF4FFE5
End If
Case 9
If Server.Switch1(LIndex).Caption = "ON" Then
Server.LEDG(8).BackColor = &HFF00&
ElseIf Server.Switch1(LIndex).Caption = "OFF" Then
Server.LEDG(8).BackColor = &HF4FFE5
End If
Case 10
If Server.Switch1(LIndex).Caption = "ON" Then
Server.LEDG(39).BackColor = &HFF00&
ElseIf Server.Switch1(LIndex).Caption = "OFF" Then
Server.LEDG(39).BackColor = &HF4FFE5
End If
Case 11
If Server.Switch1(LIndex).Caption = "ON" Then
Server.LEDG(41).BackColor = &HFF00&
ElseIf Server.Switch1(LIndex).Caption = "OFF" Then
Server.LEDG(41).BackColor = &HF4FFE5
End If
Case 12
If Server.Switch1(LIndex).Caption = "ON" Then
Server.LEDG(40).BackColor = &HFF00&
ElseIf Server.Switch1(LIndex).Caption = "OFF" Then
Server.LEDG(40).BackColor = &HF4FFE5
End If
Case 13
If Server.Switch1(LIndex).Caption = "ON" Then
Server.LEDG(43).BackColor = &HFF00&
ElseIf Server.Switch1(LIndex).Caption = "OFF" Then
Server.LEDG(43).BackColor = &HF4FFE5
End If
Case 14
If Server.Switch1(LIndex).Caption = "ON" Then
Server.LEDG(42).BackColor = &HFF00&
ElseIf Server.Switch1(LIndex).Caption = "OFF" Then
Server.LEDG(42).BackColor = &HF4FFE5
End If
Case 15
If Server.Switch1(LIndex).Caption = "ON" Then
Server.LEDG(44).BackColor = &HFF00&
ElseIf Server.Switch1(LIndex).Caption = "OFF" Then
Server.LEDG(44).BackColor = &HF4FFE5
End If
Case 16
If Server.Switch1(LIndex).Caption = "ON" Then
Server.LEDG(45).BackColor = &HFF00&
ElseIf Server.Switch1(LIndex).Caption = "OFF" Then
Server.LEDG(45).BackColor = &HF4FFE5
End If
Case 17
If Server.Switch1(LIndex).Caption = "ON" Then
Server.LEDG(46).BackColor = &HFF00&
ElseIf Server.Switch1(LIndex).Caption = "OFF" Then
Server.LEDG(46).BackColor = &HF4FFE5
End If
Case 18
If Server.Switch1(LIndex).Caption = "ON" Then
Server.LEDG(48).BackColor = &HFF00&
ElseIf Server.Switch1(LIndex).Caption = "OFF" Then
Server.LEDG(48).BackColor = &HF4FFE5
End If
Case 19
If Server.Switch1(LIndex).Caption = "ON" Then
Server.LEDG(49).BackColor = &HFF00&
ElseIf Server.Switch1(LIndex).Caption = "OFF" Then
Server.LEDG(49).BackColor = &HF4FFE5
End If
End Select
'If Server.Switch3(LIndex).Caption = "ON" Then LedFx.AllOn
'If Server.Switch3(LIndex).Caption = "OFF" Then LedFx.AllOff
End Sub
Public Sub LedFunctionsS2(LIndex As Long)
Select Case LIndex
Case 0
If Server.Switch2(LIndex).Caption = "ON" Then
Server.LEDG(65).BackColor = &HFF00&
ElseIf Server.Switch2(LIndex).Caption = "OFF" Then
Server.LEDG(65).BackColor = &HF4FFE5
End If
Case 1
If Server.Switch2(LIndex).Caption = "ON" Then
Server.LEDG(61).BackColor = &HFF00&
ElseIf Server.Switch2(LIndex).Caption = "OFF" Then
Server.LEDG(61).BackColor = &HF4FFE5
End If
Case 2
If Server.Switch2(LIndex).Caption = "ON" Then
Server.LEDG(64).BackColor = &HFF00&
ElseIf Server.Switch2(LIndex).Caption = "OFF" Then
Server.LEDG(64).BackColor = &HF4FFE5
End If
Case 3
If Server.Switch2(LIndex).Caption = "ON" Then
Server.LEDG(10).BackColor = &HFF00&
ElseIf Server.Switch2(LIndex).Caption = "OFF" Then
Server.LEDG(10).BackColor = &HF4FFE5
End If
Case 4
If Server.Switch2(LIndex).Caption = "ON" Then
Server.LEDG(9).BackColor = &HFF00&
ElseIf Server.Switch2(LIndex).Caption = "OFF" Then
Server.LEDG(9).BackColor = &HF4FFE5
End If
Case 5
If Server.Switch2(LIndex).Caption = "ON" Then
Server.LEDG(50).BackColor = &HFF00&
ElseIf Server.Switch2(LIndex).Caption = "OFF" Then
Server.LEDG(50).BackColor = &HF4FFE5
End If
Case 6
If Server.Switch2(LIndex).Caption = "ON" Then
Server.LEDG(51).BackColor = &HFF00&
ElseIf Server.Switch2(LIndex).Caption = "OFF" Then
Server.LEDG(51).BackColor = &HF4FFE5
End If
Case 7
If Server.Switch2(LIndex).Caption = "ON" Then
Server.LEDG(34).BackColor = &HFF00&
ElseIf Server.Switch2(LIndex).Caption = "OFF" Then
Server.LEDG(34).BackColor = &HF4FFE5
End If
Case 8
If Server.Switch2(LIndex).Caption = "ON" Then
Server.LEDG(35).BackColor = &HFF00&
ElseIf Server.Switch2(LIndex).Caption = "OFF" Then
Server.LEDG(35).BackColor = &HF4FFE5
End If
Case 9
If Server.Switch2(LIndex).Caption = "ON" Then
Server.LEDG(47).BackColor = &HFF00&
ElseIf Server.Switch2(LIndex).Caption = "OFF" Then
Server.LEDG(47).BackColor = &HF4FFE5
End If
End Select
'If Server.Switch3(LIndex).Caption = "ON" Then LedFx.AllOn
'If Server.Switch3(LIndex).Caption = "OFF" Then LedFx.AllOff
End Sub
Public Sub LedFunctionsS3(LIndex As Long)
Select Case LIndex
Case 0
If Server.Switch3(LIndex).Caption = "ON" Then
Server.LEDG(3).BackColor = &HFF00&
ElseIf Server.Switch3(LIndex).Caption = "OFF" Then
Server.LEDG(3).BackColor = &HF4FFE5
End If
Case 1
If Server.Switch3(LIndex).Caption = "ON" Then
Server.LEDG(4).BackColor = &HFF00&
ElseIf Server.Switch3(LIndex).Caption = "OFF" Then
Server.LEDG(4).BackColor = &HF4FFE5
End If
Case 2
If Server.Switch3(LIndex).Caption = "ON" Then
Server.LEDG(16).BackColor = &HFF00&
ElseIf Server.Switch3(LIndex).Caption = "OFF" Then
Server.LEDG(16).BackColor = &HF4FFE5
End If
Case 3
If Server.Switch3(LIndex).Caption = "ON" Then
Server.LEDG(29).BackColor = &HFF00&
ElseIf Server.Switch3(LIndex).Caption = "OFF" Then
Server.LEDG(29).BackColor = &HF4FFE5
End If
Case 4
If Server.Switch3(LIndex).Caption = "ON" Then
Server.LEDG(17).BackColor = &HFF00&
ElseIf Server.Switch3(LIndex).Caption = "OFF" Then
Server.LEDG(17).BackColor = &HF4FFE5
End If
Case 5
If Server.Switch3(LIndex).Caption = "ON" Then
Server.LEDG(18).BackColor = &HFF00&
ElseIf Server.Switch3(LIndex).Caption = "OFF" Then
Server.LEDG(18).BackColor = &HF4FFE5
End If
Case 6
If Server.Switch3(LIndex).Caption = "ON" Then
Server.LEDG(21).BackColor = &HFF00&
ElseIf Server.Switch3(LIndex).Caption = "OFF" Then
Server.LEDG(21).BackColor = &HF4FFE5
End If
Case 7
If Server.Switch3(LIndex).Caption = "ON" Then
Server.LEDG(22).BackColor = &HFF00&
ElseIf Server.Switch3(LIndex).Caption = "OFF" Then
Server.LEDG(22).BackColor = &HF4FFE5
End If
Case 8
If Server.Switch3(LIndex).Caption = "ON" Then
Server.LEDG(19).BackColor = &HFF00&
ElseIf Server.Switch3(LIndex).Caption = "OFF" Then
Server.LEDG(19).BackColor = &HF4FFE5
End If
Case 9
If Server.Switch3(LIndex).Caption = "ON" Then
Server.LEDG(19).BackColor = &HFF00&
Server.LEDG(20).BackColor = &HFF00&
ElseIf Server.Switch3(LIndex).Caption = "OFF" Then
Server.LEDG(20).BackColor = &HF4FFE5
End If
Case 10
If Server.Switch3(LIndex).Caption = "ON" Then
Server.LEDG(24).BackColor = &HFF00&
ElseIf Server.Switch3(LIndex).Caption = "OFF" Then
Server.LEDG(24).BackColor = &HF4FFE5
End If
Case 11
If Server.Switch3(LIndex).Caption = "ON" Then
Server.LEDG(23).BackColor = &HFF00&
ElseIf Server.Switch3(LIndex).Caption = "OFF" Then
Server.LEDG(23).BackColor = &HF4FFE5
End If
Case 12
If Server.Switch3(LIndex).Caption = "ON" Then
Server.LEDG(26).BackColor = &HFF00&
ElseIf Server.Switch3(LIndex).Caption = "OFF" Then
Server.LEDG(26).BackColor = &HF4FFE5
End If
Case 13
If Server.Switch3(LIndex).Caption = "ON" Then
Server.LEDG(25).BackColor = &HFF00&
ElseIf Server.Switch3(LIndex).Caption = "OFF" Then
Server.LEDG(25).BackColor = &HF4FFE5
End If
Case 14
If Server.Switch3(LIndex).Caption = "ON" Then
Server.LEDG(27).BackColor = &HFF00&
ElseIf Server.Switch3(LIndex).Caption = "OFF" Then
Server.LEDG(27).BackColor = &HF4FFE5
End If
Case 15
If Server.Switch3(LIndex).Caption = "ON" Then
Server.LEDG(28).BackColor = &HFF00&
ElseIf Server.Switch3(LIndex).Caption = "OFF" Then
Server.LEDG(28).BackColor = &HF4FFE5
End If
Case 16
Case 17
If Server.Switch3(LIndex).Caption = "ON" Then
Server.LEDG(15).BackColor = &HFF00&
ElseIf Server.Switch3(LIndex).Caption = "OFF" Then
Server.LEDG(15).BackColor = &HF4FFE5
End If
Case 18
End Select
End Sub

Public Sub AllOff()
For l = 0 To Server.LEDG.UBound
Server.LEDG(l).BackColor = &HF4FFE5
Next l
End Sub

Public Sub AllOn()
For l = 0 To Server.LEDG.UBound
Server.LEDG(l).BackColor = &HFF00&
Next l
End Sub
