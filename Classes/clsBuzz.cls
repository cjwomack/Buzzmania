VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBuzz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event ButtonPressed(Button As String, Player As Integer)

Dim Buffer(10) As Byte
Dim Player As Integer
Public Sub LightOn(Player As Integer)
    If UseBuzz = True Then OpenUSBdevice "Buzz"
    Buffer(Player) = 255
    WriteUSBdevice AddressFor(Buffer(0)), 7
    If UseBuzz = True Then CloseUSBdevice
End Sub

Public Sub LightOff(Player As Integer)
    If UseBuzz = True Then OpenUSBdevice "Buzz"
    Buffer(Player) = 0
    WriteUSBdevice AddressFor(Buffer(0)), 7
    If UseBuzz = True Then CloseUSBdevice
End Sub

Public Function GetButton(Optional Parse As Boolean = False) As Long
    modJoystick.PollJoystick
    GetButton = modJoystick.MYJOYEX.dwButtons
    If Parse Then ParseInput modJoystick.MYJOYEX.dwButtons
    If modJoystick.MYJOYEX.dwButtons > 0 Then Debug.Print "RC: " & modJoystick.MYJOYEX.dwButtons
End Function

Public Sub StartJoy()
    Dim joyNum As Integer
    Do Until modJoystick.StartJoystick(joyNum) = True Or joyNum >= 10
        joyNum = joyNum + 1
    Loop
    
End Sub

Public Sub ParseInput(TheInput As Long)
    Select Case TheInput
        Case 1
            RaiseEvent ButtonPressed("R", 1)
        Case 16
            RaiseEvent ButtonPressed("B", 1)
        Case 8
            RaiseEvent ButtonPressed("O", 1)
        Case 4
            RaiseEvent ButtonPressed("G", 1)
        Case 2
            RaiseEvent ButtonPressed("Y", 1)
            
            
        Case 32
            RaiseEvent ButtonPressed("R", 2)
        Case 64
            RaiseEvent ButtonPressed("B", 2)
        Case 128
            RaiseEvent ButtonPressed("O", 2)
        Case 256
            RaiseEvent ButtonPressed("G", 2)
        Case 512
            RaiseEvent ButtonPressed("Y", 2)
            
            
        Case 1024
            RaiseEvent ButtonPressed("R", 3)
        Case 16384
            RaiseEvent ButtonPressed("B", 3)
        Case 8192
            RaiseEvent ButtonPressed("O", 3)
        Case 4096
            RaiseEvent ButtonPressed("G", 3)
        Case 2048
            RaiseEvent ButtonPressed("Y", 3)

            
        Case 32768
            RaiseEvent ButtonPressed("R", 2)
        Case 5248288
            RaiseEvent ButtonPressed("B", 2)
        Case 262144
            RaiseEvent ButtonPressed("O", 2)
        Case 131072
            RaiseEvent ButtonPressed("G", 2)
        Case 65536
            RaiseEvent ButtonPressed("Y", 2)
    End Select
End Sub

Public Sub DoKeyboard(Key As Integer)

If Key = 122 Or Key = 90 Then
    RaiseEvent ButtonPressed("R", 1)
    Player = 1
End If

If Key = 120 Or Key = 88 Then
    RaiseEvent ButtonPressed("R", 2)
    Player = 2
End If

If Key = 99 Or Key = 67 Then
    RaiseEvent ButtonPressed("R", 3)
    Player = 3
End If

If Key = 118 Or Key = 86 Then
    RaiseEvent ButtonPressed("R", 4)
    Player = 4
End If

If Key = 49 Then RaiseEvent ButtonPressed("B", Player)
If Key = 50 Then RaiseEvent ButtonPressed("O", Player)
If Key = 51 Then RaiseEvent ButtonPressed("G", Player)
If Key = 52 Then RaiseEvent ButtonPressed("Y", Player)

End Sub

