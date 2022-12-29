VERSION 5.00
Begin VB.Form frmDecision 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6720
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10755
   LinkTopic       =   "Form1"
   ScaleHeight     =   6720
   ScaleWidth      =   10755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrWait 
      Interval        =   250
      Left            =   240
      Top             =   2760
   End
   Begin VB.Timer tmrInput 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   600
      Top             =   1440
   End
   Begin VB.Shape shpYellow 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   5
      Height          =   375
      Left            =   5760
      Shape           =   4  'Rounded Rectangle
      Top             =   6240
      Width           =   375
   End
   Begin VB.Shape shpGreen 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   5
      Height          =   375
      Left            =   4800
      Shape           =   4  'Rounded Rectangle
      Top             =   6240
      Width           =   375
   End
   Begin VB.Shape shpOrange 
      BorderColor     =   &H000080FF&
      BorderWidth     =   5
      Height          =   375
      Left            =   4200
      Shape           =   4  'Rounded Rectangle
      Top             =   6240
      Width           =   375
   End
   Begin VB.Shape shpBlue 
      BorderColor     =   &H00C0C000&
      BorderWidth     =   5
      Height          =   375
      Left            =   3600
      Shape           =   4  'Rounded Rectangle
      Top             =   6240
      Width           =   375
   End
   Begin VB.Label lblPress 
      BackStyle       =   0  'Transparent
      Caption         =   "Make your decision with      ,      ,      or        to continue.."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   6240
      Width           =   9495
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
      Left            =   1080
      TabIndex        =   4
      Top             =   2160
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
      Left            =   1080
      TabIndex        =   3
      Top             =   3000
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
      Left            =   1080
      TabIndex        =   2
      Top             =   3840
      Width           =   8535
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
      Left            =   1155
      TabIndex        =   1
      Top             =   4680
      Width           =   8535
   End
   Begin VB.Image imgYellow 
      Height          =   780
      Left            =   960
      Picture         =   "frmDecision.frx":0000
      Top             =   4560
      Width           =   8835
   End
   Begin VB.Image imgGreen 
      Height          =   780
      Left            =   960
      Picture         =   "frmDecision.frx":084E
      Top             =   3720
      Width           =   8835
   End
   Begin VB.Image imgOrange 
      Height          =   780
      Left            =   960
      Picture         =   "frmDecision.frx":1288
      Top             =   2880
      Width           =   8835
   End
   Begin VB.Image imgBlue 
      Height          =   780
      Left            =   960
      Picture         =   "frmDecision.frx":1A06
      Top             =   2040
      Width           =   8835
   End
   Begin VB.Label lblQuestion 
      BackStyle       =   0  'Transparent
      Caption         =   "Your Question goes in here.."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10335
   End
End
Attribute VB_Name = "frmDecision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents BuzzInput As clsBuzz
Attribute BuzzInput.VB_VarHelpID = -1

Private Sub BuzzInput_ButtonPressed(Button As String, Player As Integer)
    If Player = 1 Then
        If Button = "B" Then Decision = "B"
        If Button = "O" Then Decision = "O"
        If Button = "G" Then Decision = "G"
        If Button = "Y" Then Decision = "Y"
        'If Button = "R" Then Exit Sub
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Set BuzzInput = New clsBuzz
End Sub

Private Sub tmrInput_Timer()
    BuzzInput.GetButton True
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    BuzzInput.DoKeyboard (KeyAscii)
End Sub

Private Sub tmrWait_Timer()
tmrInput.Enabled = True
tmrWait.Enabled = Not tmrWait.Enabled
End Sub
