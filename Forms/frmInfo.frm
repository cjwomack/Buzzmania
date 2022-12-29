VERSION 5.00
Begin VB.Form frmInfo 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6840
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11145
   LinkTopic       =   "Form1"
   ScaleHeight     =   6840
   ScaleWidth      =   11145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrInput 
      Interval        =   100
      Left            =   4920
      Top             =   3240
   End
   Begin VB.Image imgRed 
      Height          =   555
      Left            =   1440
      Picture         =   "frmInfo.frx":0000
      Top             =   6120
      Width           =   1155
   End
   Begin VB.Label lblPress 
      BackStyle       =   0  'Transparent
      Caption         =   "Press the              button to continue"
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
      TabIndex        =   1
      Top             =   6240
      Width           =   10815
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Information goes in here.."
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
      Height          =   5655
      Left            =   1065
      TabIndex        =   0
      Top             =   120
      Width           =   9015
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents BuzzInput As clsBuzz
Attribute BuzzInput.VB_VarHelpID = -1

Private Sub BuzzInput_ButtonPressed(Button As String, Player As Integer)
    If Button = "R" Then
        Unload Me
    Else
        Exit Sub
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

