Attribute VB_Name = "modJoystick"
Option Explicit

Public Declare Function joyGetPosEx Lib "winmm.dll" (ByVal uJoyID As Long, _
pji As JOYINFOEX) As Long
Public Declare Function joyGetDevCapsA Lib "winmm.dll" (ByVal uJoyID As Long, _
pjc As JOYCAPS, ByVal cjc As Long) As Long

Public Type JOYCAPS
wMid As Integer
wPid As Integer
szPname As String * 32
wXmin As Long
wXmax As Long
wYmin As Long
wYmax As Long
wZmin As Long
wZmax As Long
wNumButtons As Long
wPeriodMin As Long
wPeriodMax As Long
wRmin As Long
wRmax As Long
wUmin As Long
wUmax As Long
wVmin As Long
wVmax As Long
wCaps As Long
wMaxAxes As Long
wNumAxes As Long
wMaxButtons As Long
szRegKey As String * 32
szOEMVxD As String * 260
End Type
Public Type JOYINFOEX
dwSize As Long
dwFlags As Long
dwXpos As Long
dwYpos As Long
dwZpos As Long
dwRpos As Long
dwUpos As Long
dwVpos As Long
dwButtons As Long
dwButtonNumber As Long
dwPOV As Long
dwReserved1 As Long
dwReserved2 As Long
End Type

Public JoyNum As Long
Public MYJOYEX As JOYINFOEX
Public MYJOYCAPS As JOYCAPS
Public CenterX As Long
Public CenterY As Long
Public JoyButtons(15) As Boolean
Public CurrentJoyX As Long
Public CurrentJoyY As Long

Public Function StartJoystick(Optional ByVal JoystickNumber As Long = 0) As Boolean
JoyNum = JoystickNumber
If joyGetDevCapsA(JoyNum, MYJOYCAPS, 404) <> 0 Then 'Get joystick info
StartJoystick = False
Else
Call joyGetPosEx(JoyNum, MYJOYEX)
CenterX = MYJOYEX.dwXpos
CenterY = MYJOYEX.dwYpos
StartJoystick = True
End If
End Function

Public Sub PollJoystick()
Dim i As Long
Dim t As Long
MYJOYEX.dwSize = 64
MYJOYEX.dwFlags = 255
' Get the joystick information
Call joyGetPosEx(JoyNum, MYJOYEX)
t = MYJOYEX.dwButtons
For i = 15 To 0 Step -1
JoyButtons(i) = False
If (2 ^ i) <= t Then
t = t - (2 ^ i)
JoyButtons(i) = True
End If
Next i
CurrentJoyX = MYJOYEX.dwXpos
CurrentJoyY = MYJOYEX.dwYpos
End Sub



