Attribute VB_Name = "ScreenFx"
Public SStyle As New MaatoohWinSets.MAppStyle
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public WinState As Boolean

Public Function ScreenFix()
Server.Width = Screen.Width '/ Screen.TwipsPerPixelX
Server.Height = Screen.Height '/ Screen.TwipsPerPixelY
Server.FrameWork.Width = (Screen.Width / Screen.TwipsPerPixelX)
Server.FrameWork.Height = (Screen.Height / Screen.TwipsPerPixelY) - 125
For k = 0 To Server.ControlFrame.UBound
Server.ControlFrame(k).Width = (Screen.Width * 0.3008)
Server.ControlFrame(k).Height = (Screen.Height * 0.4704)
Next k
For g = 0 To Server.SensorFrame.UBound
Server.SensorFrame(g).Top = Server.ControlFrame(g).Top + Server.ControlFrame(g).Height
Server.SensorFrame(g).Width = Server.ControlFrame(g).Width
Server.SensorFrame(g).Height = CLng(Screen.Height * 0.354)
Next g
Server.DriveFrame.Width = (Screen.Width * 0.02)
Call SStyle.FullContentFx(Server.hWnd, True)
On Error Resume Next
Server.EBackground.Picture = LoadPicture(App.Path & "\back.jpg")
End Function

Public Function Mark()
Server.Mark.Left = 1130
Server.Mark.Top = 750
Server.Mark.ForeColor = vbBlack
Server.Mark.Font = "MS Sans Serif"
Server.Mark.FontBold = True
Server.Mark.FontSize = 8
End Function
