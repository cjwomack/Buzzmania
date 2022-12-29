VERSION 5.00
Begin VB.Form frmPics 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7785
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8835
   LinkTopic       =   "Form1"
   ScaleHeight     =   7785
   ScaleWidth      =   8835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrTimeout 
      Interval        =   5000
      Left            =   3840
      Top             =   3600
   End
   Begin VB.PictureBox picMovies 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      Height          =   7575
      Left            =   0
      ScaleHeight     =   7515
      ScaleWidth      =   8595
      TabIndex        =   0
      Top             =   0
      Width           =   8655
   End
End
Attribute VB_Name = "frmPics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.Move Me.Left, Me.Top, Me.Width, Me.Height
End Sub

Private Sub picMovies_Paint()
Me.Move Me.Left, Me.Top, Me.Width, Me.Height

End Sub

Private Sub tmrTimeout_Timer()
    Unload Me
End Sub
