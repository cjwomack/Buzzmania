VERSION 5.00
Begin VB.Form frmMain 
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
   Begin VB.Timer tmrInput 
      Interval        =   1
      Left            =   600
      Top             =   2760
   End
   Begin VB.Label lblA4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Quit Buzzmania!"
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
      Left            =   1785
      TabIndex        =   5
      Top             =   6480
      Width           =   8535
   End
   Begin VB.Label lblA3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "About Buzzmania!"
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
      Left            =   1710
      TabIndex        =   4
      Top             =   5640
      Width           =   8535
   End
   Begin VB.Label lblA2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Change Options"
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
      Left            =   1710
      TabIndex        =   3
      Top             =   4800
      Width           =   8535
   End
   Begin VB.Label lblA1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Start Buzzmania! Game"
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
      Left            =   1710
      TabIndex        =   2
      Top             =   3960
      Width           =   8535
   End
   Begin VB.Label lblHeader 
      BackStyle       =   0  'Transparent
      Caption         =   "uzzmania 2!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1455
      Left            =   1560
      TabIndex        =   1
      Top             =   720
      Width           =   8175
   End
   Begin VB.Label lblSubText 
      BackStyle       =   0  'Transparent
      Caption         =   "The all-round Quiz game"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   615
      Left            =   2640
      TabIndex        =   0
      Top             =   2160
      Width           =   6855
   End
   Begin VB.Image imgSpeaker 
      Height          =   1965
      Left            =   120
      Picture         =   "frmMain.frx":0000
      Top             =   120
      Width           =   1515
   End
   Begin VB.Image imgBlue 
      Height          =   780
      Left            =   1590
      Picture         =   "frmMain.frx":1FF2
      Top             =   3840
      Width           =   8835
   End
   Begin VB.Image imgOrange 
      Height          =   780
      Left            =   1590
      Picture         =   "frmMain.frx":2803
      Top             =   4680
      Width           =   8835
   End
   Begin VB.Image imgGreen 
      Height          =   780
      Left            =   1590
      Picture         =   "frmMain.frx":2F81
      Top             =   5520
      Width           =   8835
   End
   Begin VB.Image imgYellow 
      Height          =   780
      Left            =   1590
      Picture         =   "frmMain.frx":39BB
      Top             =   6360
      Width           =   8835
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents BuzzInput As clsBuzz
Attribute BuzzInput.VB_VarHelpID = -1

Private Sub BuzzInput_ButtonPressed(Button As String, Player As Integer)
    If Button = "Y" Then End
    If Button = "B" Then
        tmrInput.Enabled = False
        ReDim Items(3)
        Items(0) = "1 Player"
        Items(1) = "2 Players"
        Items(2) = "3 Players"
        Items(3) = "4 Players"
        
        ShowInput ("How many players will be playing this round?")
        Select Case Decision
            Case "1 Player"
                Players(0) = True
                Players(1) = False
                Players(2) = False
                Players(3) = False
            Case "2 Players"
                Players(0) = True
                Players(1) = True
                Players(2) = False
                Players(3) = False
            Case "3 Players"
                Players(0) = True
                Players(1) = True
                Players(2) = True
                Players(3) = False
            Case "4 Players"
                Players(0) = True
                Players(1) = True
                Players(2) = True
                Players(3) = True
        End Select
        DoEvents
        
    Dim NumCategories As Integer
    SQL.dbFileName = App.Path & "\Buzz.db"
    SQL.dbOpen
    Temp = SQL.dbGetTable("SELECT COUNT(Category) FROM (SELECT DISTINCT Category FROM Questions)")
    NumCategories = Temp(1, 0)
    ReDim Items(NumCategories - 1)
    Temp = SQL.dbGetTable("SELECT DISTINCT Category FROM 'Questions'")
    For i = 1 To NumCategories
        Items(i - 1) = Temp(i, 0)
    Next i
    
    ShowInput "Please select your question category"
    LoadQuestions Decision
        
        ReDim Items(4)
        Items(0) = "30"
        Items(1) = "40"
        Items(2) = "50"
        Items(3) = "10"
        Items(4) = "20"
        ShowInput "Please select your round length"
        RoundLength = CInt(Decision)
        
        frm1P.Show
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Set BuzzInput = New clsBuzz
    UseSFX = True
    UseBuzz = True
    BuzzInput.StartJoy
    OpenUSBdevice "Buzz"
    BuzzInput.LightOn 1
    'SQL.dbGetTable "SELECT "
    ShowInfo "Welcome to Buzzmania! This new Buzz! game boasts more questions, more trivia and more game modes to keep you occupied! This message is only displayed once, and will never be displayed again on this computer. Press Red to continue, and enjoy playing Buzzmania!"
End Sub

Private Sub tmrInput_Timer()
    BuzzInput.GetButton True
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    BuzzInput.DoKeyboard (KeyAscii)
End Sub

