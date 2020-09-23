VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "A Real Card Trick"
   ClientHeight    =   2925
   ClientLeft      =   2280
   ClientTop       =   2265
   ClientWidth     =   7410
   LinkTopic       =   "Form1"
   ScaleHeight     =   2925
   ScaleWidth      =   7410
   Begin VB.CommandButton Command2 
      Caption         =   "Back"
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   2640
      Width           =   735
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   2
      Left            =   1160
      Top             =   1920
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2
      Left            =   6240
      Top             =   1920
   End
   Begin VB.CommandButton Command1 
      Caption         =   "I am done"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   2270
      Width           =   1695
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   1560
      TabIndex        =   1
      Top             =   1955
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Max             =   150
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar ProgressBar2 
      Height          =   135
      Left            =   0
      TabIndex        =   3
      Top             =   1600
      Visible         =   0   'False
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   1
      Max             =   20
   End
   Begin VB.Label Label3 
      Caption         =   "Well the secret is that pc simply changes all cards and removes only one"
      Height          =   615
      Left            =   4920
      TabIndex        =   5
      Top             =   2280
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Removing..."
      Height          =   195
      Left            =   10
      TabIndex        =   4
      Top             =   1720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Store one of these cards into your memory..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1560
      TabIndex        =   0
      Top             =   1600
      Width           =   4665
   End
   Begin VB.Image Image2 
      Height          =   1590
      Left            =   720
      Picture         =   "Form2.frx":0000
      Top             =   15
      Visible         =   0   'False
      Width           =   6165
   End
   Begin VB.Image Image1 
      Height          =   1590
      Left            =   0
      Picture         =   "Form2.frx":2000A
      Top             =   0
      Width           =   7410
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Timer1.Enabled = True
End Sub

Private Sub Command2_Click()
Main.Show
Me.Hide
End Sub

Private Sub Timer1_Timer()

If Command1.Caption = "I am done" Then

  Label1.Caption = "Reading your mind..."
  ProgressBar1.Value = ProgressBar1.Value + 1
  If ProgressBar1.Value = 130 Then Timer2.Enabled = True
    If ProgressBar1.Value = ProgressBar1.Max Then
      Command1.Caption = "Wow! Try again!"
      Timer1.Enabled = fale
      Image1.Visible = False
      Image2.Visible = True
      Label1.Caption = "''Where's my card???''"
      ProgressBar1.Value = 0
      Timer2.Enabled = False
      ProgressBar2.Value = 0
      ProgressBar2.Visible = False
      Label2.Visible = False
    End If
    
Else

   Label1.Caption = "Store one of these cards into your memory..."
   ProgressBar1.Value = 0
   Image1.Visible = True
   Image2.Visible = False
   Command1.Caption = "I am done"
   Timer1.Enabled = False
   If ProgressBar1.Max > 99 Then
      ProgressBar1.Max = ProgressBar1.Max - 50
   Else
      Exit Sub
   End If
   Label3.Visible = True
End If

End Sub

Private Sub Timer2_Timer()
  Label2.Visible = True
  ProgressBar2.Visible = True
  ProgressBar2.Value = ProgressBar2.Value + 1
End Sub
