VERSION 5.00
Begin VB.Form frm1P 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   LinkTopic       =   "Form1"
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrInputDelay 
      Interval        =   500
      Left            =   1800
      Top             =   0
   End
   Begin VB.PictureBox picTimer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      ScaleHeight     =   345
      ScaleWidth      =   4425
      TabIndex        =   6
      Top             =   1200
      Width           =   4455
   End
   Begin VB.Timer tmrWaitNext 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   600
      Top             =   0
   End
   Begin VB.Timer tmrTimeout 
      Interval        =   10000
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer tmrInput 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1200
      Top             =   0
   End
   Begin VB.Label lblA4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Answer Four"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   1380
      TabIndex        =   5
      Top             =   7200
      Width           =   8535
   End
   Begin VB.Label lblA3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Answer Three"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   1320
      TabIndex        =   4
      Top             =   6240
      Width           =   8535
   End
   Begin VB.Label lblA2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Answer Two"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   1305
      TabIndex        =   3
      Top             =   5310
      Width           =   8535
   End
   Begin VB.Label lblA1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Answer One"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   1320
      TabIndex        =   2
      Top             =   4320
      Width           =   8535
   End
   Begin VB.Label lblQuestionNo 
      BackStyle       =   0  'Transparent
      Caption         =   "Question 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   10215
   End
   Begin VB.Label lblQuestion 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   10695
   End
   Begin VB.Image Image1 
      Height          =   780
      Left            =   1200
      Picture         =   "frm1P.frx":0000
      Top             =   4200
      Width           =   8835
   End
   Begin VB.Image Image2 
      Height          =   780
      Left            =   1200
      Picture         =   "frm1P.frx":16762
      Top             =   5160
      Width           =   8835
   End
   Begin VB.Image Image3 
      Height          =   780
      Left            =   1200
      Picture         =   "frm1P.frx":2CEC4
      Top             =   6120
      Width           =   8835
   End
   Begin VB.Image Image4 
      Height          =   780
      Left            =   1200
      Picture         =   "frm1P.frx":43626
      Top             =   7080
      Width           =   8835
   End
End
Attribute VB_Name = "frm1P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' A reference to our "clsBuzz" class that lets our Multithreaded sub pass it's results
' to a class that parses our input and spits out the result
Dim WithEvents BuzzInput As clsBuzz
Attribute BuzzInput.VB_VarHelpID = -1

' This array contains our four answers. Saves having four different variables
Dim TheAnswer() As String

' Has a player buzzed in?
Dim BuzzedIn As Boolean
' Which player buzzed in?
Dim BuzzedInPlayer As Integer
' The scores for each player
Dim Score(3) As Integer
' Do we want to stop people from pressing other keys?
Dim InputLocked As Boolean
Dim i As Integer
Dim n As Integer



Private Sub Form_KeyPress(KeyAscii As Integer)
    BuzzInput.DoKeyboard (KeyAscii)
End Sub

Private Sub Form_Load()
    Set BuzzInput = New clsBuzz
    NextQuestion
End Sub


Private Sub lblQuestionNo_Click()
End
End Sub

Private Sub tmrInput_Timer()
    BuzzInput.GetButton True
    picTimer.Width = picTimer.Width - 7
End Sub

Public Sub NextQuestion()
On Error Resume Next
    If n >= RoundLength Then
        tmrInput.Enabled = False
        ShowInfo "Round over! Here's the scores: " & vbCrLf & vbCrLf & "Player 1: " & Score(0) & vbCrLf & "Player 2: " & Score(1) & vbCrLf & "Player 3: " & Score(2) & vbCrLf & "Player 4: " & Score(3) & vbCrLf & vbCrLf & "Congratulations!"
        frmMain.Show
        Unload Me
        Exit Sub
    End If
    picTimer.Width = 4455
    
    'lblQuestion.Caption = Question(i)

    ' Stops and closes the file playing with the MCI module
    modMCI.CloseFile
    
        ' Reset the buzzed in status
    BuzzedIn = False
    BuzzedInPlayer = 999
    tmrTimeout.Enabled = False
        ' Clear the image box
    'picMovies.Picture = Nothing
    Randomize
    ' Go up a question
    QuestionNumber = Int(Rnd * UBound(Question))
    If Question(QuestionNumber - 1) = "" Then NextQuestion
    n = n + 1
    ' Put the question number and the question into the respective labels
    lblQuestionNo.Caption = "Question " & n
    If QuestionNumber = 0 Then QuestionNumber = 1
    lblQuestion.Caption = Question(QuestionNumber - 1)
    
    ' Media() and Pictures() are arrays that are filled from LoadQuestions
    ' Is there media to be played? (ie. a sound file or movie), then play it
    If Media(QuestionNumber - 1) > "" Then
        frmMovies.Show
        tmrInput.Enabled = Not tmrInput.Enabled
        PlayFile Media(QuestionNumber - 1), frmMovies.picMovies
        Do Until frmMovies.Visible = False
            DoEvents
            DoEvents
        Loop
        tmrInput.Enabled = Not tmrInput.Enabled
        
    End If
    
    ' Same goes with the picture. If there's one, display it
    If Pictures(QuestionNumber - 1) > "" Then
        frmPics.Show
        tmrInput.Enabled = Not tmrInput.Enabled
        frmPics.picMovies.Picture = LoadPicture(App.Path & "\pictures\" & Pictures(QuestionNumber - 1))
        frmPics.Height = frmPics.picMovies.Height
        frmPics.Width = frmPics.picMovies.Width
        Do Until frmPics.Visible = False
            DoEvents
            DoEvents
        Loop
        tmrInput.Enabled = Not tmrInput.Enabled
        
    End If
    
        ' Turn on our 10 second rule
    tmrTimeout.Enabled = False
    tmrTimeout.Interval = 10000
    tmrTimeout.Enabled = True
    
    
    ' Read the answers from our array. They are split up with the pipe symbol
    TheAnswer() = Split(Answers(QuestionNumber - 1), "|")
    
    ' Swap the answers around to add some more fun to the game. Do this 5 times
    Randomize
    For i = 1 To 5
        SwapArrItem TheAnswer, Int(Rnd * 4), Int(Rnd * 4)
    Next i
    
    ' Set the four labels up with the answers
    lblA1.Caption = TheAnswer(0)
    lblA2.Caption = TheAnswer(1)
    lblA3.Caption = TheAnswer(2)
    lblA4.Caption = TheAnswer(3)
    Exit Sub

End Sub


Private Sub BuzzIn(Player As Integer)

' A player has buzzed in. Decide who it is, and set everything up for the answer

    ' Set the timer for 5 seconds (our buzzout time)
    tmrTimeout.Enabled = False
    tmrTimeout.Interval = 5000
    tmrTimeout.Enabled = True
    ' Set the variable to True, so nobody else can buzz in and take it all away from you
    BuzzedIn = True
    ' And also ensure this by telling the game you buzzed in
    BuzzedInPlayer = Player
    ' Stop playing the file
    modMCI.CloseFile
    ' And play the sound effect (located in ./sfx
    PlaySFX "BuzzedIn.wav"
    If UseBuzz = True Then BuzzInput.LightOn Player
    ' And ask the player to enter an answer
    lblQuestionNo.Caption = "Player " & Player & ", Answer?"
End Sub

Private Sub DoAnswer(Answer As String, Player As Integer)
    If InputLocked = True Then Exit Sub
    InputLocked = True
    
    ' If we've given the right answer?
    If Answer = Correct(QuestionNumber - 1) Then
        ' Play the ding that means we're correct
        PlaySFX "Right.wav"
        ' Add +1 to our player's score
        Score(Player - 1) = Score(Player - 1) + 1
        'lblScore(Player - 1).Caption = "P" & Player & ": " & Score(Player - 1)
    ' But if they're wrong?
    Else
        PlaySFX "Wrong.wav"
        Score(Player - 1) = Score(Player - 1) - 1
        'lblScore(Player - 1).Caption = "P" & Player & ": " & Score(Player - 1)

    End If
    
    ' Start the timer that pauses execution until we've all seen the right answer
    tmrWaitNext.Enabled = True
    tmrTimeout.Enabled = False
    ' This chunk of code hides all the wrong answers and shows the right one
    If lblA1.Caption <> Correct(QuestionNumber - 1) Then lblA1.Visible = False
    If lblA2.Caption <> Correct(QuestionNumber - 1) Then lblA2.Visible = False
    If lblA3.Caption <> Correct(QuestionNumber - 1) Then lblA3.Visible = False
    If lblA4.Caption <> Correct(QuestionNumber - 1) Then lblA4.Visible = False
    If UseBuzz = True Then BuzzInput.LightOff Player
End Sub

Private Sub tmrInputDelay_Timer()
    tmrInput.Enabled = True
End Sub

'Private Sub imgPic_KeyPress(KeyAscii As Integer)
'Form_KeyPress (KeyAscii)
'End Sub


Private Sub tmrTimeout_Timer()

    ' Start the timer that pauses execution until we've all seen the right answer
    tmrWaitNext.Enabled = True
    PlaySFX "Wrong.wav"
    ' If we've waited 5 seconds then buzz them out!
    If BuzzedInPlayer <> 999 Then DoAnswer "sdklfjsdlkf", BuzzedInPlayer

    ' This chunk of code hides all the wrong answers and shows the right one
    If lblA1.Caption <> Correct(QuestionNumber - 1) Then lblA1.Visible = False
    If lblA2.Caption <> Correct(QuestionNumber - 1) Then lblA2.Visible = False
    If lblA3.Caption <> Correct(QuestionNumber - 1) Then lblA3.Visible = False
    If lblA4.Caption <> Correct(QuestionNumber - 1) Then lblA4.Visible = False

    
    
End Sub

Private Sub tmrWaitNext_Timer()
    ' Re-show the answers and then prepare the next question
    lblA1.Visible = True
    lblA2.Visible = True
    lblA3.Visible = True
    lblA4.Visible = True
    tmrWaitNext.Enabled = False
    InputLocked = False
    tmrTimeout.Interval = 10000
    tmrTimeout.Enabled = False
    NextQuestion

End Sub

Private Sub BuzzInput_ButtonPressed(Button As String, Player As Integer)

' Think of this sub like a KeyPress sub. When you press a button on your
' Buzz! Controller, it gets parsed, and if it's a valid key, it's turned
' into something a bit more human readable ("R"ed, "B"lue, "O"range, "G"reen, "Y"ellow)
' and also shows you which controller (player) it was pressed by
    ' If nobody has buzzed in yet, then buzz them in!
    If BuzzedIn = False And Button = "R" And Players(Player - 1) = True Then
        BuzzIn Player
    End If

    ' If they have buzzed in, then pick the answer they selected
    If BuzzedIn = True And Player = BuzzedInPlayer Then
        Select Case Button
            Case "B"
                DoAnswer TheAnswer(0), Player
            Case "O"
                DoAnswer TheAnswer(1), Player
            Case "G"
                DoAnswer TheAnswer(2), Player
            Case "Y"
                DoAnswer TheAnswer(3), Player
        End Select
    End If


End Sub

