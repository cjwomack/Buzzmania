Attribute VB_Name = "modMisc"
Option Explicit
Public Items() As String

Public Decision As String
' Set this to False to not use the Buzz! Controllers. For testing purposes only
Public UseBuzz As Boolean
' A small array that lets us keep track of which players are playing
Public Players(0 To 3) As Boolean

Public RoundLength As Integer

Public Sub ShowInfo(InfoText As String)
    ' Does as it says. Shows frmInfo with a bit of information in it, then
    ' waits for the user to press the Red button before continuing
    frmInfo.Show
    frmInfo.lblInfo.Caption = InfoText
    Do Until frmInfo.Visible = False
        
        DoEvents
        DoEvents
    Loop
End Sub

Public Function ShowDecision(Question As String, Answer1 As String, Answer2 As String, Answer3 As String, Answer4 As String) As String
    ' Much the same as above, but this time lets people pick one of up to 4 decisions
    ' which is then handled by the sub that called this function.
    
    frmDecision.Show
    frmDecision.lblQuestion.Caption = Question
    
    frmDecision.lblA1.Caption = Answer1
    frmDecision.lblA2.Caption = Answer2
    frmDecision.lblA3.Caption = Answer3
    frmDecision.lblA4.Caption = Answer4
    Do Until frmDecision.Visible = False
        DoEvents
        DoEvents
    Loop
ShowDecision = Decision
End Function

Public Function ShowInput(Question As String) As String
    ' Much the same as above, but this time lets people pick one of up to 4 decisions
    ' which is then handled by the sub that called this function.
    
    frmInput.Show
    frmInput.lblQuestion.Caption = Question
    frmInput.lblTheItem.Caption = Items(0)
    frmInput.lblPrev.Caption = Items(UBound(Items))
    On Error Resume Next
    If Items(1) = "" Then frmInput.lblNext.Caption = Items(0)
    frmInput.lblNext.Caption = Items(1)

    Do Until frmInput.Visible = False
        DoEvents
        DoEvents
    Loop
ShowInput = Decision
End Function

