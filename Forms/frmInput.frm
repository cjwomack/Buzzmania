VERSION 5.00
Begin VB.Form frmInput 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10755
   LinkTopic       =   "Form1"
   ScaleHeight     =   6195
   ScaleWidth      =   10755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrInput 
      Interval        =   100
      Left            =   3360
      Top             =   3240
   End
   Begin VB.Label lblNext 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Item"
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
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   3960
      Width           =   3735
   End
   Begin VB.Label lblPrev 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Item"
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
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   3735
   End
   Begin VB.Label lblTheItem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Item"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Width           =   3735
   End
   Begin VB.Image imgRed 
      Height          =   555
      Left            =   6720
      Picture         =   "frmInput.frx":0000
      Top             =   5520
      Width           =   1155
   End
   Begin VB.Shape shpBlue 
      BorderColor     =   &H00C0C000&
      BorderWidth     =   5
      Height          =   375
      Left            =   720
      Shape           =   4  'Rounded Rectangle
      Top             =   5640
      Width           =   375
   End
   Begin VB.Shape shpOrange 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   5
      Height          =   375
      Left            =   1800
      Shape           =   4  'Rounded Rectangle
      Top             =   5640
      Width           =   375
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
   Begin VB.Label lblPress 
      BackStyle       =   0  'Transparent
      Caption         =   "Use      and      to cycle through the choices and            to select"
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
      Top             =   5640
      Width           =   9495
   End
   Begin VB.Image imgGreen 
      Height          =   780
      Left            =   -120
      Picture         =   "frmInput.frx":0A9E
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   10905
   End
End
Attribute VB_Name = "frmInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents BuzzInput As clsBuzz
Attribute BuzzInput.VB_VarHelpID = -1
Dim TheItem As Integer

Private Sub Form_KeyPress(KeyAscii As Integer)
    BuzzInput.DoKeyboard (KeyAscii)
End Sub

Private Sub BuzzInput_ButtonPressed(Button As String, Player As Integer)
    If Player = 1 And Button = "R" Then
        Decision = Items(TheItem)
        Unload Me
    End If
    
    Select Case UBound(Items)
        Case 0
            Exit Sub
        Case 1
            GoTo DoOne
        Case Else
            GoTo DoAll
    End Select

DoOne:

    If lblPrev.Caption = Items(0) Then
        lblNext.Caption = Items(1)
        lblTheItem.Caption = Items(0)
    Else
        lblNext.Caption = Items(0)
        lblTheItem.Caption = Items(1)
    End If
    
    Exit Sub

DoAll:
    
    If Player = 1 And Button = "B" Then
        TheItem = TheItem - 1
    End If
        
    If Player = 1 And Button = "Y" Then
        TheItem = TheItem + 1
    End If
    
    If TheItem = -1 Then TheItem = UBound(Items)
    If TheItem > UBound(Items) Then TheItem = 0
    
    If TheItem + 1 > UBound(Items) Then
        lblNext.Caption = Items(0)
    Else
        lblNext.Caption = Items(TheItem + 1)
    End If
    
    If TheItem - 1 = -1 Then
        lblPrev.Caption = Items(UBound(Items))
    Else
        lblPrev.Caption = Items(TheItem - 1)
    End If
    
    lblTheItem.Caption = Items(TheItem)
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Set BuzzInput = New clsBuzz

End Sub

Private Sub tmrInput_Timer()
    BuzzInput.GetButton True
End Sub

