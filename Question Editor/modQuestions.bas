Attribute VB_Name = "modQuestions"
Option Explicit
Public intQuestions As Integer
Public SQL As New clsSQL
Public Temp() As Variant
Public TotalQ As Long

Public Sub LoadQuestions()
    On Error Resume Next
    
        SQL.dbFileName = App.Path & "\Buzz.db"
    SQL.dbOpen
    
    frmMain.cmbQuestions.Clear
    intQuestions = 0
    
    Temp = SQL.dbGetTable("SELECT COUNT(ID) FROM 'Questions'")
    TotalQ = Temp(1, 0)
    Temp = SQL.dbGetTable("SELECT * FROM 'Questions'")
    
    For intQuestions = 1 To TotalQ
        frmMain.cmbQuestions.AddItem intQuestions - 1 & " - " & Temp(intQuestions, 1)
    Next intQuestions
    
    frmMain.cmbQuestions.ListIndex = 0
    

End Sub

