VERSION 5.00
Begin VB.Form frmMovies 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   6315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7890
   LinkTopic       =   "Form2"
   ScaleHeight     =   6315
   ScaleWidth      =   7890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picMovies 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6255
      Left            =   0
      ScaleHeight     =   6225
      ScaleWidth      =   7665
      TabIndex        =   0
      Top             =   0
      Width           =   7695
      Begin VB.Timer tmrTimeout 
         Interval        =   12000
         Left            =   0
         Top             =   0
      End
   End
End
Attribute VB_Name = "frmMovies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Move Me.Left, Me.Top, picMovies.Width, picMovies.Height
End Sub

Private Sub tmrTimeout_Timer()
Unload Me
End Sub
