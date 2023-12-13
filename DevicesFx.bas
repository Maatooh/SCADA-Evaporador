Attribute VB_Name = "DevicesFx"
Public TUsers As Long
Public DevName() As String
Public DevAct() As Boolean

Public Sub RegDevice(xData As String, Index As Long)
ReDim Preserve DevName(TUsers)
ReDim Preserve DevAct(TUsers)
DevName(Index) = Mid(xData, 1, InStr(xData, "- connected") - 2)
DevAct(Index) = True
'Debug.Print DevName(Index)
End Sub

Public Function PingDx(xData As String, Index As Long) As String
DevAct(Index) = True
PingDx = Replace(xData, Chr(1), vbNullString)
End Function

Public Sub delay(Timed As Double)
Dim lt As Double
lt = Timer
Do Until Timer - lt > Timed
DoEvents
Loop
End Sub
