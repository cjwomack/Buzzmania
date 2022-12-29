Attribute VB_Name = "modQuestions"
Option Explicit
' Deduct one point if they get the answer wrong?
Public DeductPoints As Boolean

Public Question() As String
Public Answers() As String
Public Media() As String
Public QuestionNumber As Integer
Public Correct() As String
Public Pictures() As String
Public SQL As New clsSQL

Public Temp() As Variant
Public i As Long


Public Sub SwapArrItem(TheArray As Variant, Item1 As Integer, Item2 As Integer)
Dim Temp As String

'hold the value of Element1 to the temporary variable
Temp = TheArray(Item1)

'set the value of Element1 equal to the value of Element 2
TheArray(Item1) = TheArray(Item2)

'set the value of Element2 equal to the value of temp (which is essencially the value of Element1)
TheArray(Item2) = Temp
End Sub

Public Sub LoadQuestions(Category As String)
    On Error Resume Next
    
    
    
    
    Dim TotalQ As Long
    
    SQL.dbFileName = App.Path & "\Buzz.db"
    SQL.dbOpen
    
    Temp = SQL.dbGetTable("SELECT COUNT(ID) FROM 'Questions' WHERE Category='" & Category & "'")
    TotalQ = Temp(1, 0)
    Temp = SQL.dbGetTable("SELECT * FROM 'Questions' WHERE Category='" & Category & "'")
    
    'MsgBox SQL.dbError
    
    For i = 1 To TotalQ
        ReDim Preserve Question(i + 1)
        ReDim Preserve Answers(i + 1)
        ReDim Preserve Media(i + 1)
        ReDim Preserve Correct(i + 1)
        ReDim Preserve Pictures(i + 1)
        Question(i) = Temp(i, 1)
        Answers(i) = Temp(i, 2)
        Correct(i) = Temp(i, 3)
        Pictures(i) = Temp(i, 5)
        Media(i) = Temp(i, 6)

    Next i
End Sub
