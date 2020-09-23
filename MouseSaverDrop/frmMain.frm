VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1470
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1950
   LinkTopic       =   "Form1"
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   1470
   ScaleWidth      =   1950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   315
      Left            =   1260
      TabIndex        =   6
      Top             =   720
      Width           =   675
   End
   Begin VB.CommandButton cmsAbout 
      Caption         =   "About"
      Height          =   315
      Left            =   660
      TabIndex        =   5
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   315
      Left            =   0
      TabIndex        =   3
      Top             =   720
      Width           =   675
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1440
      Top             =   120
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cursor Drop Program"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   1020
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "By Jason Ryczek"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label lblGrav 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Gravity:"
      BeginProperty Font 
         Name            =   "News Gothic MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label lblWind 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Wind:"
      BeginProperty Font 
         Name            =   "News Gothic MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1935
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This program can be used for whatever you like
' just please give me some credit.  Well, unless
' it's for a school project.  But, I would like
' credit if you add this to a program for something
' Thanx ~
' ~ Jason Ryczek - CCguy7@aol.com

Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Dim X1 As Integer, Y1 As Integer
Dim vmom As Integer, hmom As Integer
Dim Gravity As Integer
Dim Wind As Integer
Dim ScreenHeight As Integer, ScreenWidth As Integer

Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdStart_Click()
MsgBox "Press A Key Besides Space or Enter to Stop", , "Cursor Drop"
Timer1.Enabled = True
End Sub

Private Sub cmdStart_KeyDown(KeyCode As Integer, Shift As Integer)
Timer1.Enabled = False
End Sub

Private Sub cmsAbout_Click()
Timer1.Enabled = False
frmAbout.Show
End Sub

Private Sub Form_Load()
ScreenHeight = (Screen.Height / Screen.TwipsPerPixelY)
ScreenWidth = (Screen.Width / Screen.TwipsPerPixelX)
X1 = ScreenWidth / 2: Y1 = 0
vmom = 10
hmom = 10
Randomize Timer
Gravity = (Rnd * 8) + 2
Wind = Round((Rnd * 5) + (Rnd * -5), 1)
End Sub

Private Sub Timer1_Timer()
lblWind.Caption = "Wind: " & Wind
lblGrav.Caption = "Gravity: " & Gravity
X1 = X1 + hmom
Y1 = Y1 + vmom
' hits right side and slows down
If X1 > ScreenWidth Then X1 = ScreenWidth: hmom = -hmom * 0.9
' hits left side and slows down
If X1 < 0 Then X1 = 0: hmom = -hmom * 0.9
' hits bottom and bounces
If Y1 > ScreenHeight Then Y1 = ScreenHeight: vmom = -vmom * 0.8
' kinda floats on top
If Y1 < 0 Then Y1 = 0
'Adds gravity to the momentum
vmom = vmom + Gravity
' Adds wind to the horizontal momentum
hmom = hmom + Wind
' now, if it gets to slow, we start over
If (Abs(vmom) < 5) And (Y1 > (ScreenHeight - 10)) Then
    vmom = 1 + Rnd * 2: hmom = Round((Rnd * 2), 1) + Round((Rnd * -2), 1)
    X1 = ScreenWidth / 2: Y1 = 0
    Gravity = (Rnd * 8) + 2
    Wind = Round((Rnd * 5) + (Rnd * -5), 1)
End If
SetCursorPos X1, Y1
End Sub
