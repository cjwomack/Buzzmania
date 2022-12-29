Attribute VB_Name = "modINI"
Option Explicit

' The INI file that we will be reading from
Public INIFile As String

' The sub that lets us read from an INI file. Handy until our database is converted
' into an SQLite database ;D
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

    
' ReadText(Section to read from eg. [Section], Key to read eg. Something=value)
Public Function ReadText(Sec As String, Key As String)
Dim sRet As String
sRet = String(255, Chr(0))
ReadText = Left(sRet, GetPrivateProfileString(Sec, ByVal Key, "", sRet, 255, INIFile))
End Function
