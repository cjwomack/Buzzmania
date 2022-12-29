VERSION 5.00
Begin VB.Form frmConvert 
   Caption         =   "Convert a Buzzmania! 1 Question set"
   ClientHeight    =   1695
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   ScaleHeight     =   1695
   ScaleWidth      =   6855
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbCategory 
      Height          =   315
      ItemData        =   "frmConvert.frx":0000
      Left            =   2640
      List            =   "frmConvert.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   480
      Width           =   2895
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "Convert!"
      Height          =   375
      Left            =   5640
      TabIndex        =   2
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox txtFile 
      Height          =   285
      Left            =   2640
      TabIndex        =   1
      Text            =   "txtFile"
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label lblProgress 
      Caption         =   "Progress: 0 / 0"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label lblCategory 
      Caption         =   "Category for Questions"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label lblFile 
      Caption         =   "Buzzmania! 1 Question File (q.ini):"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmConvert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' This

Private Sub cmdConvert_Click()
'On Error Resume Next
    Dim i As Integer
    INIFile = txtFile.Text
    i = 0
    Do Until ReadText("Q" & i, "Question") = ""
        Dim Query As String
        Query = "INSERT INTO 'Questions' VALUES(null,'" & Replace(ReadText("Q" & i, "Question"), "'", Chr(34)) & "', '" & Replace(ReadText("Q" & i, "Answer"), "'", Chr(34)) & "', '" & Replace(ReadText("Q" & i, "Correct"), "'", Chr(34)) & "', '" & cmbCategory.List(cmbCategory.ListIndex) & "', '" & Replace(ReadText("Q" & i, "Image"), "'", Chr(34)) & "', '" & Replace(ReadText("Q" & i, "Media"), "'", Chr(34)) & "')"
        SQL.dbGetTable Query
        If SQL.dbError > "" Then MsgBox SQL.dbError & vbCrLf & vbCrLf & i
        lblProgress = "Progress: " & i & " questions converted"
        i = i + 1
        DoEvents
        DoEvents
    Loop

End Sub

Private Sub Form_Load()
txtFile = App.Path & "\q.ini"
cmbCategory.ListIndex = 0
End Sub
