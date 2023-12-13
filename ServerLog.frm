VERSION 5.00
Begin VB.Form ServerLog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Server - Debug Log"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5685
   Icon            =   "ServerLog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   5685
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Log 
      Height          =   2895
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   5655
   End
End
Attribute VB_Name = "ServerLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)
Cancel = True
Me.Hide
End Sub

Private Sub Log_Change()
If Len(Log) > 36000 Then
Log = vbNullString
End If
End Sub
