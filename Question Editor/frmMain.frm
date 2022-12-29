VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Buzzmania! Question Editor"
   ClientHeight    =   4410
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6525
   LinkTopic       =   "Form1"
   ScaleHeight     =   4410
   ScaleWidth      =   6525
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReload 
      Caption         =   "Reload"
      Height          =   495
      Left            =   4800
      TabIndex        =   24
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "Convert INI file"
      Height          =   495
      Left            =   120
      TabIndex        =   23
      Top             =   3720
      Width           =   1215
   End
   Begin VB.ComboBox cmbCategory 
      Height          =   315
      ItemData        =   "frmMain.frx":0000
      Left            =   960
      List            =   "frmMain.frx":0010
      TabIndex        =   21
      Text            =   "Category"
      Top             =   3000
      Width           =   4095
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear all"
      Height          =   495
      Left            =   5160
      TabIndex        =   12
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit Question"
      Height          =   495
      Left            =   2160
      TabIndex        =   9
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Select"
      Height          =   375
      Left            =   5160
      TabIndex        =   11
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save Question"
      Height          =   495
      Left            =   3480
      TabIndex        =   10
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox txtMedia 
      Height          =   285
      Left            =   960
      TabIndex        =   6
      Top             =   2280
      Width           =   4095
   End
   Begin VB.TextBox txtImage 
      Height          =   285
      Left            =   960
      TabIndex        =   7
      Top             =   2640
      Width           =   4095
   End
   Begin VB.ComboBox cmbQuestions 
      Height          =   315
      Left            =   120
      TabIndex        =   18
      Text            =   "Question 0 -"
      Top             =   120
      Width           =   4935
   End
   Begin VB.ComboBox cmbCorrect 
      Height          =   315
      Left            =   960
      TabIndex        =   8
      Text            =   "Correct Answer"
      Top             =   3360
      Width           =   4095
   End
   Begin VB.TextBox txtA4 
      Height          =   285
      Left            =   960
      TabIndex        =   5
      Top             =   1920
      Width           =   4095
   End
   Begin VB.TextBox txtA3 
      Height          =   285
      Left            =   960
      TabIndex        =   4
      Top             =   1560
      Width           =   4095
   End
   Begin VB.TextBox txtA2 
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Top             =   1200
      Width           =   4095
   End
   Begin VB.TextBox txtA1 
      Height          =   285
      Left            =   960
      TabIndex        =   2
      Top             =   840
      Width           =   4095
   End
   Begin VB.TextBox txtQuestion 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   480
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "Category:"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label lblMedia 
      Caption         =   "Media"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label lblImage 
      Caption         =   "Image"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label lblCorrect 
      Caption         =   "Correct:"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label lblA4 
      Caption         =   "Answer 4"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label lblA3 
      Caption         =   "Answer 3"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label lblA2 
      Caption         =   "Answer 2"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label lblA1 
      Caption         =   "Answer 1"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   840
      Width           =   735
   End
   Begin VB.Label lblQuestion 
      Caption         =   "Question:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   735
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' A temporary variable that holds our SQL query to be executed.
Dim Query As String

Private Sub cmdClear_Click()
        ' Does as it says, clears everything
        txtA1.Text = ""
        txtA2.Text = ""
        txtA3.Text = ""
        txtA4.Text = ""
        txtQuestion.Text = ""
        txtImage.Text = ""
        txtMedia.Text = ""
        cmbCorrect.Clear
End Sub

Private Sub cmdConvert_Click()
        ' Shows our form to convert an old Buzzmania! 1 file to a Buzzmania! 2 SQL file
        frmConvert.Show
End Sub

Private Sub cmdEdit_Click()
        ' Our UPDATE query that does as it says. It updates an existing record
        ' in our database. Questions with a single quote must be changed to
        ' a double quote (from ' to ") so it doesn't interfere with our single quotes
        Query = "UPDATE 'Questions' SET Question='" & Replace(txtQuestion.Text, "'", Chr(34)) & "', Answers='" & Replace(txtA1.Text, "'", Chr(34)) & "|" & Replace(txtA2.Text, "'", Chr(34)) & "|" & Replace(txtA3.Text, "'", Chr(34)) & "|" & Replace(txtA4.Text, "'", Chr(34)) & "', Correct='" & Replace(cmbCorrect.List(cmbCorrect.ListIndex), "'", Chr(34)) & "', Category='" & cmbCategory.List(cmbCategory.ListIndex) & "', Image='" & Replace(txtImage.Text, "'", Chr(34)) & "', Movie='" & Replace(txtMedia.Text, "'", Chr(34)) & "' WHERE ID='" & cmbQuestions.ListIndex & "'"
        ' Apply the query to the database engine
        SQL.dbGetTable Query
        
        ' Clear all the boxes
        cmdClear_Click
        
        ' Then load the questions again
        LoadQuestions
End Sub

Private Sub cmdReload_Click()
        cmdClear_Click
        LoadQuestions
End Sub

Private Sub cmdSave_Click()
        ' inserts the new question into the database. Again data is converted to prevent
        ' broken data trying to be inserted
        Query = "INSERT INTO 'Questions' VALUES(null,'" & Replace(txtQuestion.Text, "'", Chr(34)) & "', '" & Replace(txtA1.Text, "'", Chr(34)) & "|" & Replace(txtA2.Text, "'", Chr(34)) & "|" & Replace(txtA3.Text, "'", Chr(34)) & "|" & Replace(txtA4.Text, "'", Chr(34)) & "', '" & Replace(cmbCorrect.List(cmbCorrect.ListIndex), "'", Chr(34)) & "', '" & cmbCategory.List(cmbCategory.ListIndex) & "', '" & Replace(txtImage.Text, "'", Chr(34)) & "', '" & Replace(txtMedia.Text, "'", Chr(34)) & "')"
        SQL.dbGetTable Query
        cmdClear_Click
        LoadQuestions
End Sub


Private Sub cmdSelect_Click()
    ' Does as it says. Selects a question and outputs the data into textboxes
    ' and combo boxes
    
    ' If nothing's selected, quit
    If cmbQuestions.ListIndex = -1 Then Exit Sub
    ' We must offset the question number by one, as Temp(0,1) would return the column
    ' name, not the data we're after
    txtQuestion.Text = Temp(cmbQuestions.ListIndex + 1, 1)
    
    ' Answers are delimited with a pipe symbol and are split up into this array
    Dim Answers() As String
    Answers = Split(Temp(cmbQuestions.ListIndex + 1, 2), "|")
    ' Populate the text boxes with the answers
    txtA1.Text = Answers(0)
    txtA2.Text = Answers(1)
    txtA3.Text = Answers(2)
    txtA4.Text = Answers(3)
    ' Clear the correct answer box
    cmbCorrect.Clear
    ' And fill it with our 4 answers, so we can pick one from the box
    cmbCorrect.AddItem Answers(0)
    cmbCorrect.AddItem Answers(1)
    cmbCorrect.AddItem Answers(2)
    cmbCorrect.AddItem Answers(3)
    
    txtMedia.Text = Temp(cmbQuestions.ListIndex + 1, 6)
    Dim Correct As String
    Correct = Temp(cmbQuestions.ListIndex + 1, 3)
    ' Find the correct answer and set the combo box to that
    If Answers(0) = Correct Then cmbCorrect.ListIndex = 0
    If Answers(1) = Correct Then cmbCorrect.ListIndex = 1
    If Answers(2) = Correct Then cmbCorrect.ListIndex = 2
    If Answers(3) = Correct Then cmbCorrect.ListIndex = 3
    
    ' Same with the categories
    Select Case Temp(cmbQuestions.ListIndex + 1, 4)
        Case "film"
            cmbCategory.ListIndex = 0
        Case "music"
            cmbCategory.ListIndex = 1
        Case "science"
            cmbCategory.ListIndex = 2
        Case "general"
            cmbCategory.ListIndex = 3
    End Select
    
    txtImage.Text = Temp(cmbQuestions.ListIndex + 1, 5)
End Sub

Private Sub Form_Load()
    On Error Resume Next
    LoadQuestions
End Sub

Private Sub txtA1_LostFocus()
    UpdateCorrectBox
End Sub

Private Sub txtA2_LostFocus()
    UpdateCorrectBox
End Sub

Private Sub txtA3_LostFocus()
    UpdateCorrectBox
End Sub

Private Sub txtA4_LostFocus()
    UpdateCorrectBox
End Sub

Private Sub UpdateCorrectBox()
    cmbCorrect.Clear
    cmbCorrect.AddItem txtA1.Text
    cmbCorrect.AddItem txtA2.Text
    cmbCorrect.AddItem txtA3.Text
    cmbCorrect.AddItem txtA4.Text
    cmbCorrect.ListIndex = 0
End Sub
