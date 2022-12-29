VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function sqlite_open Lib "ags_sqlite.dll" (ByVal filename As String, ByVal mode As Long, ByRef errstr As String) As Long
Private Declare Sub sqlite_close Lib "ags_sqlite.dll" (ByVal DB_Handle As Long)
Private Declare Function sqlite_last_insert_rowid Lib "ags_sqlite.dll" (ByVal DB_Handle As Long) As Long
Private Declare Function sqlite_changes Lib "ags_sqlite.dll" (ByVal DB_Handle As Long) As Long
Private Declare Function sqlite_libversion Lib "ags_sqlite.dll" () As String  ' Now returns a BSTR
Private Declare Function sqlite_libencoding Lib "ags_sqlite.dll" () As String ' Now returns a BSTR
Private Declare Function ags_sqlite_get_table Lib "ags_sqlite.dll" (ByVal DB_Handle As Long, ByVal SQLString As String, ByRef errstr As String) As Variant()
'This function returns a variant SAFEARRAY otherwise it is the same as sqlite_get_table
Private Declare Function ags_sqlite_libversion Lib "ags_sqlite.dll" () As String

'This function returns the number of rows from the last sql statement.  Use this to ensure you have a valid array
Private Declare Function ags_number_of_rows_from_last_call Lib "ags_sqlite.dll" () As Long

Dim lngHandle As Long
Dim strError As String

Dim strFileName As String
Dim blnLastSucceeded As Boolean

Public Property Get dbFileName() As String
    ' Returns our filename for our SQLite database
    dbFileName = strFileName
End Property

Public Property Let dbFileName(ByVal New_File As String)
    ' Sets our SQLite database filename
    strFileName = New_File
End Property

Public Property Get dbLastSucceeded() As Boolean
    ' Did our last function succeed?
    dbLastSucceeded = blnLastSucceeded
End Property

Public Property Let dbLastSucceeded(ByVal New_Success As Boolean)
    ' Set a boolean that tells us if the last function succeeded
    blnLastSucceeded = New_Success
End Property

Public Function dbOpen() As Long
    ' Opens our database. Requires dbFileName property to be properly set.
    ' If no filename is set, then an error is raised
    lngHandle = sqlite_open(dbFileName, 0, strError)
    dbOpen = lngHandle
    dbGetHandle = lngHandle
    If strError > "" Then dbError = strError
End Function

Public Sub dbClose()
    ' As it says. Closes the connection to our SQLite database
    sqlite_close lngHandle
End Sub

Public Function dbGetTable(Optional SQLString As String = "select * from sqlite_master") As Variant
    ' Returns a variant array which contains our entire database, comma delimited
    ' (if using select * from sqlite_master), else it returns each row of results
    strError = ""
    ' Requires dbFileName to be set to a valid database file
    dbGetTable = ags_sqlite_get_table(dbGetHandle, SQLString, strError)
    
    If Trim(strError) > "" Then dbError = strError
    
    If ags_number_of_rows_from_last_call > 0 Then
        dbLastSucceeded = True
    Else
        dbLastSucceeded = False
    End If
End Function

Private Sub Class_Terminate()
    ' Cleans up our class for next time
    dbFileName = ""
    dbClose
    lngHandle = 0
End Sub

Public Property Get dbGetHandle() As Long
    dbGetHandle = lngHandle
    
End Property


Public Property Let dbGetHandle(ByVal New_Handle As Long)
    lngHandle = New_Handle
End Property


Public Property Let dbError(ByVal New_Error As String)
    strError = New_Error
End Property

Public Property Get dbError() As String
    dbError = strError
    
End Property
