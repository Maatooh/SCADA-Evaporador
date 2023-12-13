Attribute VB_Name = "SaveSensor"
Public DRCont As Long

Public Function SaveRegister()
On Error GoTo QH
If Not Dir(App.Path & "\Registers", vbDirectory) = vbNullString Then
    If Dir(App.Path & "\Registers" & "\" & Day(Date) & "-" & Month(Date) & "-" & Year(Date), vbDirectory) = vbNullString Then
    MkDir (App.Path & "\Registers" & "\" & Day(Date) & "-" & Month(Date) & "-" & Year(Date))
    End If
End If
'--Comparative
Dim TReg(23) As String
For s = 0 To UBound(TReg)
TReg(s) = Server.Sensor(s)
Next s
Call SaveComparative("Registros.csv", TReg)
QH:
End Function

Public Function SaveComparative(NameComparative As String, ArrayData() As String)
On Error GoTo QH
'--Create New
If Dir(App.Path & "\Registers" & "\" & Day(Date) & "-" & Month(Date) & "-" & Year(Date) & "\" & NameComparative, vbArchive) = vbNullString Then
'---Labels---
Dim LLabel As String
For l = 0 To 23
LLabel = LLabel & Server.LSENSOR(l) & ";"
Next l
'-------------
Open App.Path & "\Registers" & "\" & Day(Date) & "-" & Month(Date) & "-" & Year(Date) & "\" & NameComparative For Output As #1
Print #1, "Hora;" & LLabel
Close #1
End If
'--Add Data
Dim KData As String
For k = LBound(ArrayData) To UBound(ArrayData)
KData = KData & ArrayData(k) & ";"
Next k
KData = Mid(KData, 1, Len(KData) - 1)
If Not Dir(App.Path & "\Registers" & "\" & Day(Date) & "-" & Month(Date) & "-" & Year(Date) & "\" & NameComparative, vbArchive) = vbNullString Then
Open App.Path & "\Registers" & "\" & Day(Date) & "-" & Month(Date) & "-" & Year(Date) & "\" & NameComparative For Append As #1
Print #1, Time & ";" & KData
Close #1
End If
QH:
End Function




