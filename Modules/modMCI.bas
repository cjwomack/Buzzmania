Attribute VB_Name = "modMCI"
Option Explicit

Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Private Const SND_FILENAME = &H20000     ' name is a file name
Private Const SND_ASYNC = &H1

Public UseSFX As Boolean

Public Sub OpenFile(File As String)
    CloseFile
    mciSendString "open " & File & " alias MyMedia", 0, 0, 0
End Sub

Public Sub CloseFile()
    mciSendString "stop MyMedia", 0, 0, 0
    mciSendString "close MyMedia", 0, 0, 0
End Sub

Public Sub SetVideoLocation(Destination As String)
    'Set the image playback of the media file to the picture box
    mciSendString "window MyMedia handle " & Destination, 0, 0, 0

    'Set the image so it will be resized to match the picture box
    mciSendString "put MyMedia destination", 0, 0, 0

End Sub

Public Sub Play()
    mciSendString "play MyMedia", 0, 0, 0
End Sub

Public Sub Stop1()
    mciSendString "stop MyMedia", 0, 0, 0
End Sub

Public Sub PlayFile(Filename As String, Container As PictureBox)
    modMCI.OpenFile App.Path & "\Media\" & Filename
    modMCI.SetVideoLocation Str(Container.hWnd)
    modMCI.Play
    
End Sub

Public Sub PlaySFX(Filename As String)
If UseSFX = True Then PlaySound App.Path & "\sfx\" & Filename, ByVal 0&, SND_FILENAME Or SND_ASYNC
End Sub

Public Function GetPlayBackStatus() As String
'Declare a fixed length string to store the data returned by the
'mciSendString
Dim retVal As String * 15

'Open the media file
mciSendString "open c:\windows\clock.avi alias clock", 0, 0, 0

'Start playing the file
mciSendString "play clock", 0, 0, 0

'Get the current playback status of the file. This status will
'be stored in the string retVal
mciSendString "status clock mode", retVal, 15, 0
GetPlayBackStatus = retVal
End Function
