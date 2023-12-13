VERSION 5.00
Begin VB.Form Devices 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Check Devices"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3045
   Icon            =   "Devices.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   3045
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton SIP4 
      Caption         =   "Show Ipv4"
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   3360
      Width           =   3015
   End
   Begin VB.Frame Frame1 
      Caption         =   "Devices Connected"
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3015
      Begin VB.ListBox DeviceList 
         Height          =   2985
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   2775
      End
   End
End
Attribute VB_Name = "Devices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Unload(Cancel As Integer)
Cancel = True
Hide
End Sub

Private Sub SIP4_Click()
Dim xDev As String
For X = 0 To Server.WServer.UBound
If Server.WServer(X).State = sckConnected Then
xDev = xDev & Server.WServer(X).RemoteHostIP & vbCrLf
'Server.WServer(X).Close
End If
Next X
MsgBox xDev, vbInformation + vbOKOnly, "Maltexco - Devices Ipv4"
End Sub
