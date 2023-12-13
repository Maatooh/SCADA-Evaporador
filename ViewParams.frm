VERSION 5.00
Begin VB.Form ViewParams 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Establecer Parámetros"
   ClientHeight    =   720
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3105
   Icon            =   "ViewParams.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   720
   ScaleWidth      =   3105
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer ClickPS 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2880
      Top             =   0
   End
   Begin VB.CommandButton SetValue 
      Caption         =   "Establecer"
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox Param 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Text            =   "0"
      Top             =   360
      Width           =   1335
   End
   Begin VB.CommandButton Plus 
      Caption         =   "+"
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   3
      Top             =   360
      Width           =   375
   End
   Begin VB.CommandButton Plus 
      Caption         =   "-"
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   375
   End
   Begin VB.Label TitleParam 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3135
   End
End
Attribute VB_Name = "ViewParams"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private ClickPos As Boolean
Private ClickAct As Boolean
Private LastIndex As Long

Public Function LoadParam(Index As Long)
Me.Show
TitleParam.Caption = Server.LSENSOR(Index).Caption
Param = ControlFx.SensorParams(Index)
LastIndex = Index
Param.SelStart = Len(Param)
End Function

Private Sub ClickPS_Timer()
Select Case ClickPos
Case True
Param = Val(Param) + 1
Case False
Param = Val(Param) - 1
End Select
End Sub

Private Sub Param_Change()
If Len(Param) = 1 Then
Param.SelStart = 1
End If
If Not IsNumeric(Param) = True And InStr(Param, "-") = 0 Then
Param = 0
Else
'---Limit----
If Val(Param) > 100 Then
Param = 100
End If
If Val(Param) < 0 Then
Param = 0
End If
'------------
Param = Val(Param)
End If
End Sub

Private Sub Param_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyUp
Param = Val(Param) + 1
SetValue_Click
Case vbKeyDown
Param = Val(Param) - 1
SetValue_Click
Case vbKeyReturn
SetValue_Click
End Select
End Sub

Private Sub Param_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ClickPS.Enabled = False
End Sub

Private Sub Plus_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'--------
Select Case Index
Case 1
ClickPos = True
Param = Val(Param) + 1
Case 0
ClickPos = False
Param = Val(Param) - 1
End Select
ClickPS.Enabled = True
ClickAct = True
'-------
Dim TimeL As Long
TimeL = Timer
Do Until Timer - TimeL > 0.8
If ClickAct = False Then
ClickPS.Enabled = False
Exit Sub
End If
DoEvents
Loop
'ClickPS.Interval = 100
Select Case Index
Case 1
ClickPos = True
ClickPS.Enabled = True
Case 0
ClickPos = False
ClickPS.Enabled = True
End Select
End Sub

Private Sub Plus_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Do Until ClickPS.Enabled = False
ClickPS.Enabled = False
ClickAct = False
SetValue_Click
DoEvents
Loop
End Sub

Public Sub SetValue_Click()
ControlFx.SensorParams(LastIndex) = Param
Call ControlFx.SetPID("P0", "PDev0")
'---Save----------
SaveFx.SaveSwitch
'-----------------
End Sub

Private Sub SetValue_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ClickPS.Enabled = False
End Sub

Private Sub TitleParam_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ClickPS.Enabled = False
End Sub
